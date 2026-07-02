"""Experiment 2: non-parametric Bayes, Fisher linear classifier, and LOO."""

from __future__ import annotations

import csv
import math
from pathlib import Path
from typing import Dict, List, Sequence, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.metrics import accuracy_score, confusion_matrix


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "experiment" / "男女data数据集" / "data"
OUT_DIR = ROOT / "results" / "experiment2_linear"
FIG_DIR = ROOT / "reports" / "figures" / "experiment2"

LABELS = ("F", "M")
PRIOR_EQUAL = (0.5, 0.5)


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


def fit_gaussian(x: np.ndarray, y: np.ndarray) -> Dict[str, Dict[str, np.ndarray]]:
    model = {}
    for label in LABELS:
        samples = x[y == label]
        mean = samples.mean(axis=0)
        centered = samples - mean
        cov = centered.T @ centered / len(samples)
        cov = cov + np.eye(cov.shape[0]) * 1e-6
        model[label] = {"mean": mean, "cov": cov}
    return model


def log_gaussian_density(x: np.ndarray, mean: np.ndarray, cov: np.ndarray) -> np.ndarray:
    inv_cov = np.linalg.inv(cov)
    sign, logdet = np.linalg.slogdet(cov)
    if sign <= 0:
        raise ValueError("Covariance matrix is not positive definite.")
    diff = x - mean
    quad = np.sum((diff @ inv_cov) * diff, axis=1)
    return -0.5 * (x.shape[1] * math.log(2 * math.pi) + logdet + quad)


def predict_gaussian(model: Dict[str, Dict[str, np.ndarray]], x: np.ndarray, priors: Tuple[float, float]) -> np.ndarray:
    scores = []
    for label, prior in zip(LABELS, priors):
        params = model[label]
        scores.append(log_gaussian_density(x, params["mean"], params["cov"]) + math.log(prior))
    score_matrix = np.vstack(scores).T
    return np.array([LABELS[idx] for idx in np.argmax(score_matrix, axis=1)])


def parzen_density(query: np.ndarray, samples: np.ndarray, bandwidth: float) -> np.ndarray:
    diff = query[:, None, :] - samples[None, :, :]
    exponent = -0.5 * np.sum((diff / bandwidth) ** 2, axis=2)
    coef = 1.0 / ((2 * math.pi) ** (samples.shape[1] / 2) * bandwidth ** samples.shape[1])
    return coef * np.exp(exponent).mean(axis=1)


def predict_parzen(
    train_x: np.ndarray,
    train_y: np.ndarray,
    query: np.ndarray,
    bandwidth: float,
    priors: Tuple[float, float] = PRIOR_EQUAL,
) -> np.ndarray:
    scores = []
    for label, prior in zip(LABELS, priors):
        samples = train_x[train_y == label]
        density = parzen_density(query, samples, bandwidth)
        scores.append(np.log(density + 1e-300) + math.log(prior))
    score_matrix = np.vstack(scores).T
    return np.array([LABELS[idx] for idx in np.argmax(score_matrix, axis=1)])


def fit_fisher(x: np.ndarray, y: np.ndarray) -> Dict[str, np.ndarray]:
    f = x[y == "F"]
    m = x[y == "M"]
    mean_f = f.mean(axis=0)
    mean_m = m.mean(axis=0)
    sw = (f - mean_f).T @ (f - mean_f) + (m - mean_m).T @ (m - mean_m)
    w = np.linalg.pinv(sw) @ (mean_m - mean_f)
    proj_f = f @ w
    proj_m = m @ w
    threshold = 0.5 * (proj_f.mean() + proj_m.mean())
    direction = 1.0 if proj_m.mean() > proj_f.mean() else -1.0
    return {
        "w": w,
        "threshold": np.array([threshold]),
        "direction": np.array([direction]),
        "mean_f": mean_f,
        "mean_m": mean_m,
    }


def predict_fisher(model: Dict[str, np.ndarray], x: np.ndarray) -> np.ndarray:
    scores = x @ model["w"]
    threshold = float(model["threshold"][0])
    direction = float(model["direction"][0])
    is_m = scores >= threshold if direction > 0 else scores < threshold
    return np.where(is_m, "M", "F")


def loo_error_parzen(x: np.ndarray, y: np.ndarray, bandwidth: float) -> float:
    errors = 0
    for idx in range(len(x)):
        mask = np.ones(len(x), dtype=bool)
        mask[idx] = False
        pred = predict_parzen(x[mask], y[mask], x[idx : idx + 1], bandwidth)[0]
        errors += int(pred != y[idx])
    return errors / len(x)


def loo_error_fisher(x: np.ndarray, y: np.ndarray) -> float:
    errors = 0
    for idx in range(len(x)):
        mask = np.ones(len(x), dtype=bool)
        mask[idx] = False
        model = fit_fisher(x[mask], y[mask])
        pred = predict_fisher(model, x[idx : idx + 1])[0]
        errors += int(pred != y[idx])
    return errors / len(x)


def loo_error_gaussian(x: np.ndarray, y: np.ndarray) -> float:
    errors = 0
    for idx in range(len(x)):
        mask = np.ones(len(x), dtype=bool)
        mask[idx] = False
        model = fit_gaussian(x[mask], y[mask])
        pred = predict_gaussian(model, x[idx : idx + 1], PRIOR_EQUAL)[0]
        errors += int(pred != y[idx])
    return errors / len(x)


def metrics_row(method: str, setting: str, split: str, y_true: np.ndarray, y_pred: np.ndarray) -> Dict[str, object]:
    cm = confusion_matrix(y_true, y_pred, labels=list(LABELS))
    return {
        "method": method,
        "setting": setting,
        "split": split,
        "accuracy": accuracy_score(y_true, y_pred),
        "error_rate": 1 - accuracy_score(y_true, y_pred),
        "cm_FF": int(cm[0, 0]),
        "cm_FM": int(cm[0, 1]),
        "cm_MF": int(cm[1, 0]),
        "cm_MM": int(cm[1, 1]),
    }


def write_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


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


def format_percent(value: float) -> str:
    return f"{value * 100:.2f}\\%"


def write_summary_table(path: Path, rows: List[Dict[str, object]], loo: Dict[str, float]) -> None:
    lookup = {(row["method"], row["split"]): row for row in rows}
    methods = [
        ("Gaussian Bayes", "参数高斯 Bayes"),
        ("Parzen Bayes", "Parzen 窗 Bayes"),
        ("Fisher", "Fisher 线性判别"),
    ]
    lines = [
        "\\begin{tabular}{p{3.0cm}p{2.0cm}p{2.0cm}p{2.0cm}p{2.0cm}}",
        "\\toprule",
        "方法 & 训练错误率 & test1错误率 & test2错误率 & 留一法错误率 \\\\",
        "\\midrule",
    ]
    for method, label in methods:
        train = lookup[(method, "train")]
        test1 = lookup[(method, "test1")]
        test2 = lookup[(method, "test2")]
        lines.append(
            f"{latex_escape(label)} & {format_percent(train['error_rate'])} & "
            f"{format_percent(test1['error_rate'])} & {format_percent(test2['error_rate'])} & "
            f"{format_percent(loo[method])} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def plot_boundaries(
    train_x: np.ndarray,
    train_y: np.ndarray,
    gaussian_model: Dict[str, Dict[str, np.ndarray]],
    fisher_model: Dict[str, np.ndarray],
) -> None:
    x_min, x_max = train_x[:, 0].min() - 6, train_x[:, 0].max() + 6
    y_min, y_max = train_x[:, 1].min() - 8, train_x[:, 1].max() + 8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 300), np.linspace(y_min, y_max, 300))
    grid = np.c_[xx.ravel(), yy.ravel()]
    gaussian_pred = predict_gaussian(gaussian_model, grid, PRIOR_EQUAL)
    zz = (gaussian_pred == "M").astype(int).reshape(xx.shape)

    plt.figure(figsize=(7, 5))
    plt.contour(xx, yy, zz, levels=[0.5], colors="#984ea3", linewidths=1.5)
    w = fisher_model["w"]
    threshold = float(fisher_model["threshold"][0])
    x_line = np.linspace(x_min, x_max, 100)
    if abs(w[1]) > 1e-12:
        y_line = (threshold - w[0] * x_line) / w[1]
        plt.plot(x_line, y_line, color="#ff7f00", linewidth=1.6, label="Fisher boundary")
    for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
        mask = train_y == label
        plt.scatter(train_x[mask, 0], train_x[mask, 1], marker=marker, color=color, label=f"Train {label}", alpha=0.78)
    plt.plot([], [], color="#984ea3", linewidth=1.5, label="Bayes boundary")
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title("Fisher and Bayes Decision Boundaries")
    plt.legend()
    plt.grid(alpha=0.22)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "fisher_bayes_boundaries.pdf")
    plt.close()


def plot_parzen_bayes_boundaries(
    train_x: np.ndarray,
    train_y: np.ndarray,
    gaussian_model: Dict[str, Dict[str, np.ndarray]],
    bandwidth: float,
) -> None:
    x_min, x_max = train_x[:, 0].min() - 6, train_x[:, 0].max() + 6
    y_min, y_max = train_x[:, 1].min() - 8, train_x[:, 1].max() + 8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 260), np.linspace(y_min, y_max, 260))
    grid = np.c_[xx.ravel(), yy.ravel()]
    methods = [
        ("Gaussian Bayes", predict_gaussian(gaussian_model, grid, PRIOR_EQUAL)),
        (f"Parzen Bayes (h={bandwidth})", predict_parzen(train_x, train_y, grid, bandwidth)),
    ]

    fig, axes = plt.subplots(1, 2, figsize=(10, 4.2), sharex=True, sharey=True)
    for ax, (title, pred) in zip(axes, methods):
        zz = (pred == "M").astype(int).reshape(xx.shape)
        ax.contourf(xx, yy, zz, levels=[-0.5, 0.5, 1.5], colors=["#fde6d7", "#d7efec"], alpha=0.85)
        ax.contour(xx, yy, zz, levels=[0.5], colors="black", linewidths=0.9)
        for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
            mask = train_y == label
            ax.scatter(
                train_x[mask, 0],
                train_x[mask, 1],
                marker=marker,
                color=color,
                edgecolor="white",
                linewidth=0.25,
                s=18,
                alpha=0.82,
                label=label,
            )
        ax.set_title(title)
        ax.set_xlabel("Height")
        ax.grid(alpha=0.18)
    axes[0].set_ylabel("Weight")
    axes[1].legend(loc="lower right", frameon=True)
    fig.suptitle("Parametric and Non-parametric Bayes Boundaries")
    fig.tight_layout()
    fig.savefig(FIG_DIR / "parzen_bayes_boundaries.pdf")
    plt.close(fig)


def plot_projection(train_x: np.ndarray, train_y: np.ndarray, fisher_model: Dict[str, np.ndarray]) -> None:
    projection = train_x @ fisher_model["w"]
    threshold = float(fisher_model["threshold"][0])
    plt.figure(figsize=(7, 3.8))
    for label, color in [("F", "#d95f02"), ("M", "#1b9e77")]:
        values = projection[train_y == label]
        plt.hist(values, bins=12, alpha=0.58, color=color, label=label)
    plt.axvline(threshold, color="black", linestyle="--", label="threshold")
    plt.xlabel("Fisher projection")
    plt.ylabel("Count")
    plt.title("Training Samples on Fisher Projection")
    plt.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "fisher_projection.pdf")
    plt.close()


def plot_error_comparison(rows: List[Dict[str, object]], loo: Dict[str, float]) -> None:
    methods = ["Gaussian Bayes", "Parzen Bayes", "Fisher"]
    labels = ["Gaussian\nBayes", "Parzen\nBayes", "Fisher"]
    lookup = {(row["method"], row["split"]): float(row["error_rate"]) for row in rows}
    train_errors = [lookup[(method, "train")] for method in methods]
    test2_errors = [lookup[(method, "test2")] for method in methods]
    loo_errors = [loo[method] for method in methods]

    x = np.arange(len(methods))
    width = 0.25
    plt.figure(figsize=(7.2, 4.2))
    plt.bar(x - width, train_errors, width, label="Train error", color="#8da0cb")
    plt.bar(x, loo_errors, width, label="LOO error", color="#fc8d62")
    plt.bar(x + width, test2_errors, width, label="test2 error", color="#66c2a5")
    plt.xticks(x, labels)
    plt.ylabel("Error rate")
    plt.ylim(0, max(loo_errors + train_errors + test2_errors) + 0.05)
    plt.title("Training, Leave-one-out, and Test Error")
    plt.legend()
    plt.grid(axis="y", alpha=0.22)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "loo_error_comparison.pdf")
    plt.close()


def main() -> None:
    ensure_dirs()
    data = load_data()
    train_x, train_y = data["train"]

    gaussian_model = fit_gaussian(train_x, train_y)
    fisher_model = fit_fisher(train_x, train_y)
    bandwidth = 5.0

    rows: List[Dict[str, object]] = []
    for split, (x_split, y_split) in data.items():
        pred_gaussian = predict_gaussian(gaussian_model, x_split, PRIOR_EQUAL)
        pred_parzen = predict_parzen(train_x, train_y, x_split, bandwidth)
        pred_fisher = predict_fisher(fisher_model, x_split)
        rows.append(metrics_row("Gaussian Bayes", "full covariance; equal prior", split, y_split, pred_gaussian))
        rows.append(metrics_row("Parzen Bayes", f"Gaussian kernel h={bandwidth}; equal prior", split, y_split, pred_parzen))
        rows.append(metrics_row("Fisher", "height and weight", split, y_split, pred_fisher))

    loo = {
        "Gaussian Bayes": loo_error_gaussian(train_x, train_y),
        "Parzen Bayes": loo_error_parzen(train_x, train_y, bandwidth),
        "Fisher": loo_error_fisher(train_x, train_y),
    }

    write_csv(OUT_DIR / "metrics.csv", rows)
    write_summary_table(OUT_DIR / "summary_table.tex", rows, loo)
    plot_parzen_bayes_boundaries(train_x, train_y, gaussian_model, bandwidth)
    plot_boundaries(train_x, train_y, gaussian_model, fisher_model)
    plot_projection(train_x, train_y, fisher_model)
    plot_error_comparison(rows, loo)

    with (OUT_DIR / "fisher_params.txt").open("w", encoding="utf-8") as f:
        f.write(f"Fisher w: {fisher_model['w'].round(8).tolist()}\n")
        f.write(f"Fisher threshold: {float(fisher_model['threshold'][0]):.8f}\n")
        f.write(f"LOO Gaussian Bayes error: {loo['Gaussian Bayes']:.6f}\n")
        f.write(f"LOO Parzen Bayes error: {loo['Parzen Bayes']:.6f}\n")
        f.write(f"LOO Fisher error: {loo['Fisher']:.6f}\n")

    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
