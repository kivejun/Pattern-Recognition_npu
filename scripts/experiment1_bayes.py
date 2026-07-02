"""Experiment 1: Bayes gender classification with height/weight data."""

from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Dict, Iterable, List, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.metrics import accuracy_score, confusion_matrix, precision_score, recall_score


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "experiment" / "男女data数据集" / "data"
OUT_DIR = ROOT / "results" / "experiment1_bayes"
FIG_DIR = ROOT / "reports" / "figures" / "experiment1"

LABELS = ("F", "M")
PRIORS = [(0.5, 0.5), (0.75, 0.25), (0.9, 0.1)]


def ensure_dirs() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FIG_DIR.mkdir(parents=True, exist_ok=True)


def normalize_label(label: str) -> str:
    value = label.strip().upper()
    if value.startswith("F"):
        return "F"
    if value.startswith("M"):
        return "M"
    raise ValueError(f"Unknown label: {label}")


def read_train_file(path: Path, label: str) -> Tuple[np.ndarray, np.ndarray]:
    rows = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            height, weight = line.split()[:2]
            rows.append((float(height), float(weight)))
    return np.array(rows, dtype=float), np.array([label] * len(rows))


def read_test_file(path: Path) -> Tuple[np.ndarray, np.ndarray]:
    rows = []
    labels = []
    with path.open("r", encoding="utf-8") as f:
        for line in f:
            if not line.strip():
                continue
            height, weight, label = line.split()[:3]
            rows.append((float(height), float(weight)))
            labels.append(normalize_label(label))
    return np.array(rows, dtype=float), np.array(labels)


def load_data() -> Dict[str, Tuple[np.ndarray, np.ndarray]]:
    female_x, female_y = read_train_file(DATA_DIR / "FEMALE.TXT", "F")
    male_x, male_y = read_train_file(DATA_DIR / "MALE.TXT", "M")
    train_x = np.vstack([female_x, male_x])
    train_y = np.concatenate([female_y, male_y])
    test1_x, test1_y = read_test_file(DATA_DIR / "test1.txt")
    test2_x, test2_y = read_test_file(DATA_DIR / "test2.txt")
    return {
        "train": (train_x, train_y),
        "test1": (test1_x, test1_y),
        "test2": (test2_x, test2_y),
    }


def fit_gaussian(
    x: np.ndarray,
    y: np.ndarray,
    features: Sequence[int],
    covariance_mode: str,
) -> Dict[str, Dict[str, np.ndarray]]:
    model: Dict[str, Dict[str, np.ndarray]] = {}
    for label in LABELS:
        class_x = x[y == label][:, features]
        if class_x.ndim == 1:
            class_x = class_x[:, None]
        mean = class_x.mean(axis=0)
        centered = class_x - mean
        cov = centered.T @ centered / len(class_x)
        cov = np.atleast_2d(cov)
        if covariance_mode == "diag":
            cov = np.diag(np.diag(cov))
        cov = cov + np.eye(cov.shape[0]) * 1e-6
        model[label] = {"mean": mean, "cov": cov}
    return model


def log_gaussian_density(x: np.ndarray, mean: np.ndarray, cov: np.ndarray) -> np.ndarray:
    x = np.atleast_2d(x)
    dim = len(mean)
    inv_cov = np.linalg.inv(cov)
    sign, logdet = np.linalg.slogdet(cov)
    if sign <= 0:
        raise ValueError("Covariance matrix is not positive definite.")
    diff = x - mean
    quad = np.sum((diff @ inv_cov) * diff, axis=1)
    return -0.5 * (dim * math.log(2 * math.pi) + logdet + quad)


def predict_gaussian(
    model: Dict[str, Dict[str, np.ndarray]],
    x: np.ndarray,
    features: Sequence[int],
    priors: Tuple[float, float],
) -> np.ndarray:
    x_selected = x[:, features]
    if x_selected.ndim == 1:
        x_selected = x_selected[:, None]
    scores = []
    for label, prior in zip(LABELS, priors):
        params = model[label]
        scores.append(log_gaussian_density(x_selected, params["mean"], params["cov"]) + math.log(prior))
    score_matrix = np.vstack(scores).T
    return np.array([LABELS[idx] for idx in np.argmax(score_matrix, axis=1)])


def predict_min_risk(
    model: Dict[str, Dict[str, np.ndarray]],
    x: np.ndarray,
    features: Sequence[int],
    priors: Tuple[float, float],
    loss_predict_f_when_m: float = 1.0,
    loss_predict_m_when_f: float = 2.0,
) -> np.ndarray:
    x_selected = x[:, features]
    log_posts = []
    for label, prior in zip(LABELS, priors):
        params = model[label]
        log_posts.append(log_gaussian_density(x_selected, params["mean"], params["cov"]) + math.log(prior))
    log_posts = np.vstack(log_posts).T
    max_log = np.max(log_posts, axis=1, keepdims=True)
    posts = np.exp(log_posts - max_log)
    posts = posts / posts.sum(axis=1, keepdims=True)
    posterior_f = posts[:, 0]
    posterior_m = posts[:, 1]
    risk_predict_f = loss_predict_f_when_m * posterior_m
    risk_predict_m = loss_predict_m_when_f * posterior_f
    return np.where(risk_predict_f <= risk_predict_m, "F", "M")


def parzen_density(
    query: np.ndarray,
    samples: np.ndarray,
    bandwidth: float,
) -> np.ndarray:
    query = np.atleast_2d(query)
    dim = samples.shape[1]
    diff = query[:, None, :] - samples[None, :, :]
    exponent = -0.5 * np.sum((diff / bandwidth) ** 2, axis=2)
    coef = 1.0 / ((2 * math.pi) ** (dim / 2) * bandwidth**dim)
    return coef * np.exp(exponent).mean(axis=1)


def predict_parzen(
    train_x: np.ndarray,
    train_y: np.ndarray,
    x: np.ndarray,
    features: Sequence[int],
    priors: Tuple[float, float],
    bandwidth: float,
) -> np.ndarray:
    x_selected = x[:, features]
    class_scores = []
    for label, prior in zip(LABELS, priors):
        samples = train_x[train_y == label][:, features]
        density = parzen_density(x_selected, samples, bandwidth)
        class_scores.append(np.log(density + 1e-300) + math.log(prior))
    scores = np.vstack(class_scores).T
    return np.array([LABELS[idx] for idx in np.argmax(scores, axis=1)])


def predict_knn(
    train_x: np.ndarray,
    train_y: np.ndarray,
    x: np.ndarray,
    features: Sequence[int],
    k: int,
) -> np.ndarray:
    train_selected = train_x[:, features]
    query = x[:, features]
    preds = []
    for row in query:
        distances = np.linalg.norm(train_selected - row, axis=1)
        nearest = np.argsort(distances)[:k]
        labels, counts = np.unique(train_y[nearest], return_counts=True)
        max_count = counts.max()
        tied = sorted(labels[counts == max_count])
        preds.append(tied[0])
    return np.array(preds)


def metrics_row(method: str, setting: str, split: str, y_true: np.ndarray, y_pred: np.ndarray) -> Dict[str, object]:
    cm = confusion_matrix(y_true, y_pred, labels=list(LABELS))
    return {
        "method": method,
        "setting": setting,
        "split": split,
        "accuracy": accuracy_score(y_true, y_pred),
        "precision_F": precision_score(y_true, y_pred, labels=list(LABELS), pos_label="F", zero_division=0),
        "recall_F": recall_score(y_true, y_pred, labels=list(LABELS), pos_label="F", zero_division=0),
        "precision_M": precision_score(y_true, y_pred, labels=list(LABELS), pos_label="M", zero_division=0),
        "recall_M": recall_score(y_true, y_pred, labels=list(LABELS), pos_label="M", zero_division=0),
        "cm_FF": int(cm[0, 0]),
        "cm_FM": int(cm[0, 1]),
        "cm_MF": int(cm[1, 0]),
        "cm_MM": int(cm[1, 1]),
    }


def write_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    if not rows:
        return
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def format_percent(value: float) -> str:
    return f"{value * 100:.2f}\\%"


def latex_escape(value: object) -> str:
    text = str(value)
    replacements = {
        "\\": "\\textbackslash{}",
        "_": "\\_",
        "%": "\\%",
        "&": "\\&",
        "#": "\\#",
        "{": "\\{",
        "}": "\\}",
    }
    return "".join(replacements.get(ch, ch) for ch in text)


def write_latex_table(path: Path, rows: List[Dict[str, object]], selected_settings: Sequence[Tuple[str, str, str]]) -> None:
    lookup = {(row["method"], row["setting"], row["split"]): row for row in rows}
    lines = [
        "\\begin{tabular}{p{3.2cm}p{4.0cm}p{1.8cm}p{1.8cm}p{2.4cm}}",
        "\\toprule",
        "方法 & 设置 & 训练准确率 & test1准确率 & test2准确率 \\\\",
        "\\midrule",
    ]
    for method, setting, label in selected_settings:
        train = lookup[(method, setting, "train")]
        test1 = lookup[(method, setting, "test1")]
        test2 = lookup[(method, setting, "test2")]
        lines.append(
            f"{latex_escape(label)} & {latex_escape(setting)} & {format_percent(train['accuracy'])} & "
            f"{format_percent(test1['accuracy'])} & {format_percent(test2['accuracy'])} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def plot_data(train_x: np.ndarray, train_y: np.ndarray, test1_x: np.ndarray, test1_y: np.ndarray) -> None:
    plt.figure(figsize=(7, 5))
    for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
        mask = train_y == label
        plt.scatter(train_x[mask, 0], train_x[mask, 1], marker=marker, color=color, label=f"Train {label}", alpha=0.75)
    for label, marker, color in [("F", "x", "#e7298a"), ("M", "+", "#377eb8")]:
        mask = test1_y == label
        plt.scatter(test1_x[mask, 0], test1_x[mask, 1], marker=marker, color=color, label=f"Test1 {label}", alpha=0.85)
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title("Gender Dataset: Height and Weight")
    plt.legend()
    plt.grid(alpha=0.25)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "data_scatter.pdf")
    plt.close()


def plot_decision_boundary(model: Dict[str, Dict[str, np.ndarray]], train_x: np.ndarray, train_y: np.ndarray) -> None:
    x_min, x_max = train_x[:, 0].min() - 6, train_x[:, 0].max() + 6
    y_min, y_max = train_x[:, 1].min() - 8, train_x[:, 1].max() + 8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 240), np.linspace(y_min, y_max, 240))
    grid = np.c_[xx.ravel(), yy.ravel()]
    preds = predict_gaussian(model, grid, [0, 1], (0.5, 0.5))
    zz = (preds == "M").astype(int).reshape(xx.shape)
    plt.figure(figsize=(7, 5))
    plt.contourf(xx, yy, zz, levels=[-0.5, 0.5, 1.5], colors=["#fddbc7", "#c7eae5"], alpha=0.55)
    plt.contour(xx, yy, zz, levels=[0.5], colors="black", linewidths=1.2)
    for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
        mask = train_y == label
        plt.scatter(train_x[mask, 0], train_x[mask, 1], marker=marker, color=color, label=f"Train {label}", edgecolor="white")
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title("2D Gaussian Bayes Decision Boundary")
    plt.legend()
    plt.grid(alpha=0.2)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "bayes_2d_boundary.pdf")
    plt.close()


def plot_confusion(cm: np.ndarray, title: str, path: Path) -> None:
    fig, ax = plt.subplots(figsize=(4.2, 3.6))
    im = ax.imshow(cm, cmap="Blues")
    ax.set_xticks([0, 1], LABELS)
    ax.set_yticks([0, 1], LABELS)
    ax.set_xlabel("Predicted")
    ax.set_ylabel("True")
    ax.set_title(title)
    for i in range(2):
        for j in range(2):
            ax.text(j, i, str(cm[i, j]), ha="center", va="center", color="black")
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    plt.tight_layout()
    plt.savefig(path)
    plt.close()


def plot_1d_gaussian_density(
    train_x: np.ndarray,
    train_y: np.ndarray,
    feature_idx: int,
    feature_name: str,
    path: Path,
) -> None:
    """Draw one-dimensional class densities estimated by MLE."""
    x_min = train_x[:, feature_idx].min() - 8
    x_max = train_x[:, feature_idx].max() + 8
    grid = np.linspace(x_min, x_max, 500)
    plt.figure(figsize=(7, 4.2))
    for label, color, text in [("F", "#d95f02", "Female"), ("M", "#1b9e77", "Male")]:
        values = train_x[train_y == label, feature_idx]
        mean = values.mean()
        var = ((values - mean) ** 2).mean()
        density = np.exp(-0.5 * (grid - mean) ** 2 / var) / math.sqrt(2 * math.pi * var)
        plt.plot(grid, density, color=color, linewidth=2.0, label=f"{text}: mean={mean:.2f}, std={math.sqrt(var):.2f}")
        plt.hist(values, bins=10, density=True, alpha=0.18, color=color)
    plt.xlabel(feature_name)
    plt.ylabel("Density")
    plt.title(f"MLE Gaussian Density by {feature_name}")
    plt.grid(alpha=0.25)
    plt.legend()
    plt.tight_layout()
    plt.savefig(path)
    plt.close()


def plot_prior_boundaries(model: Dict[str, Dict[str, np.ndarray]], train_x: np.ndarray, train_y: np.ndarray) -> None:
    """Compare how prior probabilities move the 2D Bayes decision boundary."""
    x_min, x_max = train_x[:, 0].min() - 6, train_x[:, 0].max() + 6
    y_min, y_max = train_x[:, 1].min() - 8, train_x[:, 1].max() + 8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 260), np.linspace(y_min, y_max, 260))
    grid = np.c_[xx.ravel(), yy.ravel()]
    plt.figure(figsize=(7, 5))
    for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
        mask = train_y == label
        plt.scatter(train_x[mask, 0], train_x[mask, 1], marker=marker, color=color, label=f"Train {label}", alpha=0.62)
    styles = [((0.5, 0.5), "black", "-"), ((0.75, 0.25), "#7570b3", "--"), ((0.9, 0.1), "#e7298a", "-.")]
    for priors, color, linestyle in styles:
        preds = predict_gaussian(model, grid, [0, 1], priors)
        zz = (preds == "M").astype(int).reshape(xx.shape)
        plt.contour(xx, yy, zz, levels=[0.5], colors=color, linewidths=1.8, linestyles=linestyle)
    plt.plot([], [], color="black", linestyle="-", label="P(F)=0.50")
    plt.plot([], [], color="#7570b3", linestyle="--", label="P(F)=0.75")
    plt.plot([], [], color="#e7298a", linestyle="-.", label="P(F)=0.90")
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title("Effect of Priors on Bayes Boundary")
    plt.grid(alpha=0.2)
    plt.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "bayes_prior_boundaries.pdf")
    plt.close()


def plot_nonparametric_boundaries(train_x: np.ndarray, train_y: np.ndarray) -> None:
    """Draw Parzen and kNN decision regions for the two-feature experiment."""
    x_min, x_max = train_x[:, 0].min() - 6, train_x[:, 0].max() + 6
    y_min, y_max = train_x[:, 1].min() - 8, train_x[:, 1].max() + 8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 220), np.linspace(y_min, y_max, 220))
    grid = np.c_[xx.ravel(), yy.ravel()]
    fig, axes = plt.subplots(1, 2, figsize=(10, 4.2), sharex=True, sharey=True)
    methods = [
        ("Parzen Bayes (h=5.0)", predict_parzen(train_x, train_y, grid, [0, 1], (0.5, 0.5), bandwidth=5.0)),
        ("kNN (k=5)", predict_knn(train_x, train_y, grid, [0, 1], k=5)),
    ]
    for ax, (title, preds) in zip(axes, methods):
        zz = (preds == "M").astype(int).reshape(xx.shape)
        ax.contourf(xx, yy, zz, levels=[-0.5, 0.5, 1.5], colors=["#fddbc7", "#c7eae5"], alpha=0.55)
        ax.contour(xx, yy, zz, levels=[0.5], colors="black", linewidths=1.0)
        for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
            mask = train_y == label
            ax.scatter(train_x[mask, 0], train_x[mask, 1], marker=marker, color=color, alpha=0.75, edgecolor="white")
        ax.set_title(title)
        ax.set_xlabel("Height")
        ax.grid(alpha=0.18)
    axes[0].set_ylabel("Weight")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "nonparametric_boundaries.pdf")
    plt.close()


def main() -> None:
    ensure_dirs()
    data = load_data()
    train_x, train_y = data["train"]
    test1_x, test1_y = data["test1"]
    test2_x, test2_y = data["test2"]

    rows: List[Dict[str, object]] = []
    split_data = {"train": (train_x, train_y), "test1": (test1_x, test1_y), "test2": (test2_x, test2_y)}

    feature_settings = [
        ("height", [0], "full"),
        ("weight", [1], "full"),
        ("height_weight_full_cov", [0, 1], "full"),
        ("height_weight_diag_cov", [0, 1], "diag"),
    ]
    for feature_name, features, cov_mode in feature_settings:
        model = fit_gaussian(train_x, train_y, features, cov_mode)
        for priors in PRIORS:
            setting = f"{feature_name}; P(F)={priors[0]:.2f}, P(M)={priors[1]:.2f}"
            for split, (x_split, y_split) in split_data.items():
                pred = predict_gaussian(model, x_split, features, priors)
                rows.append(metrics_row("Gaussian Bayes", setting, split, y_split, pred))

    full_model = fit_gaussian(train_x, train_y, [0, 1], "full")
    risk_setting = "height_weight_full_cov; P(F)=0.50, P(M)=0.50; loss M|F=2, loss F|M=1"
    for split, (x_split, y_split) in split_data.items():
        pred = predict_min_risk(full_model, x_split, [0, 1], (0.5, 0.5))
        rows.append(metrics_row("Minimum Risk Bayes", risk_setting, split, y_split, pred))

    parzen_setting = "height_weight; Gaussian kernel h=5.0; P(F)=0.50, P(M)=0.50"
    for split, (x_split, y_split) in split_data.items():
        pred = predict_parzen(train_x, train_y, x_split, [0, 1], (0.5, 0.5), bandwidth=5.0)
        rows.append(metrics_row("Parzen Bayes", parzen_setting, split, y_split, pred))

    knn_setting = "height_weight; k=5"
    for split, (x_split, y_split) in split_data.items():
        pred = predict_knn(train_x, train_y, x_split, [0, 1], k=5)
        rows.append(metrics_row("kNN", knn_setting, split, y_split, pred))

    write_csv(OUT_DIR / "metrics.csv", rows)
    selected = [
        ("Gaussian Bayes", "height; P(F)=0.50, P(M)=0.50", "单特征 Bayes"),
        ("Gaussian Bayes", "weight; P(F)=0.50, P(M)=0.50", "单特征 Bayes"),
        ("Gaussian Bayes", "height_weight_full_cov; P(F)=0.50, P(M)=0.50", "二维相关 Bayes"),
        ("Gaussian Bayes", "height_weight_diag_cov; P(F)=0.50, P(M)=0.50", "二维不相关 Bayes"),
        ("Gaussian Bayes", "height_weight_full_cov; P(F)=0.75, P(M)=0.25", "不同先验 Bayes"),
        ("Gaussian Bayes", "height_weight_full_cov; P(F)=0.90, P(M)=0.10", "不同先验 Bayes"),
        ("Minimum Risk Bayes", risk_setting, "最小风险 Bayes"),
        ("Parzen Bayes", parzen_setting, "Parzen Bayes"),
        ("kNN", knn_setting, "kNN"),
    ]
    write_latex_table(OUT_DIR / "summary_table.tex", rows, selected)

    plot_data(train_x, train_y, test1_x, test1_y)
    plot_1d_gaussian_density(train_x, train_y, 0, "Height", FIG_DIR / "single_feature_height_density.pdf")
    plot_1d_gaussian_density(train_x, train_y, 1, "Weight", FIG_DIR / "single_feature_weight_density.pdf")
    plot_decision_boundary(full_model, train_x, train_y)
    plot_prior_boundaries(full_model, train_x, train_y)
    plot_nonparametric_boundaries(train_x, train_y)
    best_pred_test2 = predict_gaussian(full_model, test2_x, [0, 1], (0.5, 0.5))
    best_cm = confusion_matrix(test2_y, best_pred_test2, labels=list(LABELS))
    plot_confusion(best_cm, "Gaussian Bayes on Test2", FIG_DIR / "bayes_test2_confusion.pdf")

    with (OUT_DIR / "decision_rules.txt").open("w", encoding="utf-8") as f:
        for label in LABELS:
            params = full_model[label]
            f.write(f"{label} mean: {params['mean'].round(6).tolist()}\n")
            f.write(f"{label} covariance: {params['cov'].round(6).tolist()}\n")
        f.write("\nDecision rule for minimum-error Bayes with equal prior:\n")
        f.write("choose class argmax_i log N(x; mu_i, Sigma_i) + log P(omega_i).\n")
        f.write("\nMinimum-risk loss table used in this experiment:\n")
        f.write("lambda(F|F)=0, lambda(M|M)=0, lambda(F|M)=1, lambda(M|F)=2.\n")

    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
