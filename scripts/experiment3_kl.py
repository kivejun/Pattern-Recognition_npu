"""Experiment 3: K-L transform/PCA feature extraction for gender data."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.metrics import accuracy_score, confusion_matrix


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "experiment" / "男女data数据集" / "data"
OUT_DIR = ROOT / "results" / "experiment3_kl"
FIG_DIR = ROOT / "reports" / "figures" / "experiment3"

LABELS = ("F", "M")


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


def fit_pca_kl(x: np.ndarray) -> Dict[str, np.ndarray]:
    mean = x.mean(axis=0)
    centered = x - mean
    cov = centered.T @ centered / len(x)
    eigvals, eigvecs = np.linalg.eigh(cov)
    order = np.argsort(eigvals)[::-1]
    eigvals = eigvals[order]
    eigvecs = eigvecs[:, order]
    if eigvecs[0, 0] < 0:
        eigvecs[:, 0] *= -1
    return {"mean": mean, "cov": cov, "eigvals": eigvals, "eigvecs": eigvecs}


def fit_mean_direction(x: np.ndarray, y: np.ndarray) -> Dict[str, np.ndarray]:
    mean_f = x[y == "F"].mean(axis=0)
    mean_m = x[y == "M"].mean(axis=0)
    direction = mean_m - mean_f
    direction = direction / np.linalg.norm(direction)
    midpoint = 0.5 * (mean_f + mean_m)
    threshold = float(midpoint @ direction)
    return {
        "direction": direction,
        "threshold": np.array([threshold]),
        "mean_f": mean_f,
        "mean_m": mean_m,
    }


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
    return {"direction": w / np.linalg.norm(w), "w": w, "threshold": np.array([threshold])}


def fit_projection_classifier(x: np.ndarray, y: np.ndarray, direction: np.ndarray) -> Dict[str, np.ndarray]:
    proj = x @ direction
    mean_f = proj[y == "F"].mean()
    mean_m = proj[y == "M"].mean()
    threshold = 0.5 * (mean_f + mean_m)
    sign = 1.0 if mean_m > mean_f else -1.0
    return {"direction": direction, "threshold": np.array([threshold]), "sign": np.array([sign])}


def predict_projection(model: Dict[str, np.ndarray], x: np.ndarray) -> np.ndarray:
    scores = x @ model["direction"]
    threshold = float(model["threshold"][0])
    sign = float(model.get("sign", np.array([1.0]))[0])
    is_m = scores >= threshold if sign > 0 else scores < threshold
    return np.where(is_m, "M", "F")


def predict_fisher(model: Dict[str, np.ndarray], x: np.ndarray) -> np.ndarray:
    scores = x @ model["w"]
    threshold = float(model["threshold"][0])
    return np.where(scores >= threshold, "M", "F")


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
    return -0.5 * (x.shape[1] * np.log(2 * np.pi) + logdet + quad)


def predict_gaussian(model: Dict[str, Dict[str, np.ndarray]], x: np.ndarray) -> np.ndarray:
    scores = []
    for label in LABELS:
        params = model[label]
        scores.append(log_gaussian_density(x, params["mean"], params["cov"]) + np.log(0.5))
    score_matrix = np.vstack(scores).T
    return np.array([LABELS[idx] for idx in np.argmax(score_matrix, axis=1)])


def metrics_row(method: str, split: str, y_true: np.ndarray, y_pred: np.ndarray) -> Dict[str, object]:
    cm = confusion_matrix(y_true, y_pred, labels=list(LABELS))
    accuracy = accuracy_score(y_true, y_pred)
    return {
        "method": method,
        "split": split,
        "accuracy": accuracy,
        "error_rate": 1 - accuracy,
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


def format_percent(value: float) -> str:
    return f"{value * 100:.2f}\\%"


def write_summary_table(path: Path, rows: List[Dict[str, object]]) -> None:
    lookup = {(row["method"], row["split"]): row for row in rows}
    methods = [
        ("K-L PCA 主成分", "K-L/PCA 主方向"),
        ("类均值方向", "类均值差方向"),
        ("Fisher", "Fisher 线性判别"),
        ("Gaussian Bayes", "二维高斯 Bayes"),
    ]
    lines = [
        "\\begin{tabular}{p{3.0cm}p{2.2cm}p{2.2cm}p{2.2cm}}",
        "\\toprule",
        "方法 & 训练准确率 & test1准确率 & test2准确率 \\\\",
        "\\midrule",
    ]
    for method, label in methods:
        train = lookup[(method, "train")]
        test1 = lookup[(method, "test1")]
        test2 = lookup[(method, "test2")]
        lines.append(
            f"{label} & {format_percent(train['accuracy'])} & "
            f"{format_percent(test1['accuracy'])} & {format_percent(test2['accuracy'])} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def draw_direction(ax, center: np.ndarray, direction: np.ndarray, scale: float, color: str, label: str) -> None:
    start = center - direction * scale
    end = center + direction * scale
    ax.annotate(
        "",
        xy=end,
        xytext=start,
        arrowprops={"arrowstyle": "->", "color": color, "linewidth": 2.0},
    )
    ax.text(end[0], end[1], label, color=color, fontsize=10)


def plot_directions(
    train_x: np.ndarray,
    train_y: np.ndarray,
    pca: Dict[str, np.ndarray],
    mean_model: Dict[str, np.ndarray],
    fisher_model: Dict[str, np.ndarray],
) -> None:
    fig, ax = plt.subplots(figsize=(7, 5))
    for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
        mask = train_y == label
        ax.scatter(train_x[mask, 0], train_x[mask, 1], marker=marker, color=color, label=label, alpha=0.78)
    center = pca["mean"]
    draw_direction(ax, center, pca["eigvecs"][:, 0], 18, "#7570b3", "PC1")
    draw_direction(ax, center, pca["eigvecs"][:, 1], 12, "#1f78b4", "PC2")
    draw_direction(ax, center, mean_model["direction"], 16, "#e7298a", "Mean diff")
    draw_direction(ax, center, fisher_model["direction"], 16, "#ff7f00", "Fisher")
    ax.set_xlabel("Height")
    ax.set_ylabel("Weight")
    ax.set_title("K-L Directions and Discriminant Directions")
    ax.legend()
    ax.grid(alpha=0.22)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "kl_directions.pdf")
    plt.close()


def plot_projection(
    train_x: np.ndarray,
    train_y: np.ndarray,
    model: Dict[str, np.ndarray],
    title: str,
    path: Path,
) -> None:
    projection = train_x @ model["direction"]
    threshold = float(model["threshold"][0])
    plt.figure(figsize=(7, 3.8))
    for label, color in [("F", "#d95f02"), ("M", "#1b9e77")]:
        values = projection[train_y == label]
        plt.hist(values, bins=12, alpha=0.58, color=color, label=label)
    plt.axvline(threshold, color="black", linestyle="--", label="threshold")
    plt.xlabel("Projection value")
    plt.ylabel("Count")
    plt.title(title)
    plt.legend()
    plt.tight_layout()
    plt.savefig(path)
    plt.close()


def plot_classification_boundaries(
    train_x: np.ndarray,
    train_y: np.ndarray,
    pca_model: Dict[str, np.ndarray],
    mean_model: Dict[str, np.ndarray],
    fisher_model: Dict[str, np.ndarray],
    gaussian_model: Dict[str, Dict[str, np.ndarray]],
) -> None:
    x_min, x_max = train_x[:, 0].min() - 6, train_x[:, 0].max() + 6
    y_min, y_max = train_x[:, 1].min() - 8, train_x[:, 1].max() + 8
    xx, yy = np.meshgrid(np.linspace(x_min, x_max, 280), np.linspace(y_min, y_max, 280))
    grid = np.c_[xx.ravel(), yy.ravel()]

    bayes_pred = predict_gaussian(gaussian_model, grid)
    bayes_region = (bayes_pred == "M").astype(int).reshape(xx.shape)

    plt.figure(figsize=(7.6, 5.2))
    plt.contourf(xx, yy, bayes_region, levels=[-0.5, 0.5, 1.5], colors=["#fde6d7", "#d7efec"], alpha=0.45)
    plt.contour(xx, yy, bayes_region, levels=[0.5], colors="#984ea3", linewidths=1.5)

    line_specs = [
        (pca_model["direction"], float(pca_model["threshold"][0]), "#7570b3", "PCA/PC1 boundary"),
        (mean_model["direction"], float(mean_model["threshold"][0]), "#e7298a", "Mean-diff boundary"),
        (fisher_model["w"], float(fisher_model["threshold"][0]), "#ff7f00", "Fisher boundary"),
    ]
    x_line = np.linspace(x_min, x_max, 200)
    for direction, threshold, color, label in line_specs:
        if abs(direction[1]) > 1e-12:
            y_line = (threshold - direction[0] * x_line) / direction[1]
            plt.plot(x_line, y_line, color=color, linewidth=1.6, label=label)

    for label, marker, color in [("F", "o", "#d95f02"), ("M", "^", "#1b9e77")]:
        mask = train_y == label
        plt.scatter(
            train_x[mask, 0],
            train_x[mask, 1],
            marker=marker,
            color=color,
            edgecolor="white",
            linewidth=0.25,
            s=22,
            label=f"Train {label}",
            alpha=0.82,
        )
    plt.plot([], [], color="#984ea3", linewidth=1.5, label="Gaussian Bayes boundary")
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title("K-L, Mean Difference, Fisher, and Bayes Boundaries")
    plt.legend(fontsize=8)
    plt.grid(alpha=0.22)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "classification_boundaries.pdf")
    plt.close()


def main() -> None:
    ensure_dirs()
    data = load_data()
    train_x, train_y = data["train"]

    pca = fit_pca_kl(train_x)
    pca_model = fit_projection_classifier(train_x, train_y, pca["eigvecs"][:, 0])
    mean_model = fit_mean_direction(train_x, train_y)
    mean_model["sign"] = np.array([1.0])
    fisher_model = fit_fisher(train_x, train_y)
    gaussian_model = fit_gaussian(train_x, train_y)

    rows: List[Dict[str, object]] = []
    for split, (x_split, y_split) in data.items():
        rows.append(metrics_row("K-L PCA 主成分", split, y_split, predict_projection(pca_model, x_split)))
        rows.append(metrics_row("类均值方向", split, y_split, predict_projection(mean_model, x_split)))
        rows.append(metrics_row("Fisher", split, y_split, predict_fisher(fisher_model, x_split)))
        rows.append(metrics_row("Gaussian Bayes", split, y_split, predict_gaussian(gaussian_model, x_split)))

    write_csv(OUT_DIR / "metrics.csv", rows)
    write_summary_table(OUT_DIR / "summary_table.tex", rows)

    plot_directions(train_x, train_y, pca, mean_model, fisher_model)
    plot_projection(train_x, train_y, pca_model, "Projection on First K-L/PCA Component", FIG_DIR / "pca_projection.pdf")
    plot_projection(train_x, train_y, mean_model, "Projection on Class-Mean Difference Direction", FIG_DIR / "mean_projection.pdf")
    plot_classification_boundaries(train_x, train_y, pca_model, mean_model, fisher_model, gaussian_model)

    explained = pca["eigvals"] / pca["eigvals"].sum()
    with (OUT_DIR / "kl_params.txt").open("w", encoding="utf-8") as f:
        f.write(f"Global mean: {pca['mean'].round(8).tolist()}\n")
        f.write(f"Covariance: {pca['cov'].round(8).tolist()}\n")
        f.write(f"Eigenvalues: {pca['eigvals'].round(8).tolist()}\n")
        f.write(f"Explained ratios: {explained.round(8).tolist()}\n")
        f.write(f"PC1: {pca['eigvecs'][:, 0].round(8).tolist()}\n")
        f.write(f"PC2: {pca['eigvecs'][:, 1].round(8).tolist()}\n")
        f.write(f"Class mean direction: {mean_model['direction'].round(8).tolist()}\n")
        f.write(f"Class mean threshold: {float(mean_model['threshold'][0]):.8f}\n")

    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
