"""Experiment 5: C-means/k-means and hierarchical clustering."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Tuple

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from scipy.cluster.hierarchy import dendrogram, linkage
from sklearn.cluster import AgglomerativeClustering
from sklearn.metrics import adjusted_rand_score, confusion_matrix, silhouette_score
from sklearn.preprocessing import StandardScaler


ROOT = Path(__file__).resolve().parents[1]
DATA_DIR = ROOT / "experiment" / "男女data数据集" / "data"
OUT_DIR = ROOT / "results" / "experiment5_clustering"
FIG_DIR = ROOT / "reports" / "figures" / "experiment5"

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
    test2_x, test2_y = read_test_file(DATA_DIR / "test2.txt")
    combined_x = np.vstack([train_x, test2_x])
    combined_y = np.concatenate([train_y, test2_y])
    return {
        "train": (train_x, train_y),
        "test2": (test2_x, test2_y),
        "combined": (combined_x, combined_y),
    }


def c_means(
    x: np.ndarray,
    n_clusters: int,
    seed: int,
    max_iter: int = 200,
    tol: float = 1e-6,
) -> Tuple[np.ndarray, np.ndarray, float, int]:
    rng = np.random.default_rng(seed)
    indices = rng.choice(len(x), size=n_clusters, replace=False)
    centers = x[indices].copy()
    labels = np.zeros(len(x), dtype=int)
    for iteration in range(1, max_iter + 1):
        distances = np.linalg.norm(x[:, None, :] - centers[None, :, :], axis=2)
        new_labels = distances.argmin(axis=1)
        new_centers = centers.copy()
        for cluster in range(n_clusters):
            members = x[new_labels == cluster]
            if len(members) == 0:
                new_centers[cluster] = x[rng.integers(0, len(x))]
            else:
                new_centers[cluster] = members.mean(axis=0)
        shift = np.linalg.norm(new_centers - centers)
        centers = new_centers
        labels = new_labels
        if shift < tol:
            break
    inertia = float(np.sum((x - centers[labels]) ** 2))
    return labels, centers, inertia, iteration


def best_label_accuracy(y_true: np.ndarray, cluster_labels: np.ndarray) -> float:
    true_binary = np.array([0 if label == "F" else 1 for label in y_true])
    if len(np.unique(cluster_labels)) != 2:
        return float("nan")
    acc_a = np.mean(cluster_labels == true_binary)
    acc_b = np.mean((1 - cluster_labels) == true_binary)
    return float(max(acc_a, acc_b))


def cluster_to_gender_table(y_true: np.ndarray, cluster_labels: np.ndarray) -> np.ndarray:
    true_binary = np.array([0 if label == "F" else 1 for label in y_true])
    if len(np.unique(cluster_labels)) != 2:
        return np.zeros((2, 2), dtype=int)
    cm_a = confusion_matrix(true_binary, cluster_labels, labels=[0, 1])
    cm_b = confusion_matrix(true_binary, 1 - cluster_labels, labels=[0, 1])
    return cm_a if np.trace(cm_a) >= np.trace(cm_b) else cm_b


def evaluate_k_range(x_scaled: np.ndarray, y: np.ndarray, dataset_name: str) -> List[Dict[str, object]]:
    rows: List[Dict[str, object]] = []
    for n_clusters in range(2, 6):
        labels, centers, inertia, iterations = c_means(x_scaled, n_clusters, seed=2026 + n_clusters)
        silhouette = silhouette_score(x_scaled, labels)
        ari = adjusted_rand_score(y, labels)
        rows.append(
            {
                "dataset": dataset_name,
                "method": "C-means",
                "n_clusters": n_clusters,
                "seed": 2026 + n_clusters,
                "inertia": inertia,
                "silhouette": silhouette,
                "ari": ari,
                "iterations": iterations,
                "gender_accuracy_if_2": best_label_accuracy(y, labels) if n_clusters == 2 else "",
            }
        )
    return rows


def evaluate_initial_seeds(x_scaled: np.ndarray, y: np.ndarray) -> List[Dict[str, object]]:
    rows = []
    for seed in range(10):
        labels, centers, inertia, iterations = c_means(x_scaled, 2, seed=seed)
        rows.append(
            {
                "dataset": "train",
                "method": "C-means different init",
                "n_clusters": 2,
                "seed": seed,
                "inertia": inertia,
                "silhouette": silhouette_score(x_scaled, labels),
                "ari": adjusted_rand_score(y, labels),
                "iterations": iterations,
                "gender_accuracy_if_2": best_label_accuracy(y, labels),
            }
        )
    return rows


def evaluate_hierarchical(x_scaled: np.ndarray, y: np.ndarray, dataset_name: str) -> Dict[str, object]:
    clustering = AgglomerativeClustering(n_clusters=2, linkage="ward")
    labels = clustering.fit_predict(x_scaled)
    return {
        "dataset": dataset_name,
        "method": "Hierarchical Ward",
        "n_clusters": 2,
        "seed": "",
        "inertia": "",
        "silhouette": silhouette_score(x_scaled, labels),
        "ari": adjusted_rand_score(y, labels),
        "iterations": "",
        "gender_accuracy_if_2": best_label_accuracy(y, labels),
    }


def write_csv(path: Path, rows: List[Dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def format_float(value: object, digits: int = 4) -> str:
    if value == "":
        return ""
    return f"{float(value):.{digits}f}"


def format_percent(value: object) -> str:
    if value == "":
        return ""
    return f"{float(value) * 100:.2f}\\%"


def write_summary_table(path: Path, rows: List[Dict[str, object]]) -> None:
    selected = [
        row for row in rows
        if (row["dataset"], row["method"], int(row["n_clusters"])) in {
            ("train", "C-means", 2),
            ("train", "Hierarchical Ward", 2),
            ("combined", "C-means", 2),
            ("combined", "Hierarchical Ward", 2),
        }
    ]
    lines = [
        "\\begin{tabular}{p{2.4cm}p{3.0cm}p{1.4cm}p{1.8cm}p{1.8cm}p{2.0cm}}",
        "\\toprule",
        "数据集 & 方法 & 类数 & 轮廓系数 & ARI & 性别匹配率 \\\\",
        "\\midrule",
    ]
    name_map = {"train": "训练集", "combined": "训练集+test2"}
    method_map = {"C-means": "C 均值", "Hierarchical Ward": "Ward 分级聚类"}
    for row in selected:
        lines.append(
            f"{name_map[row['dataset']]} & {method_map[row['method']]} & {int(row['n_clusters'])} & "
            f"{format_float(row['silhouette'])} & {format_float(row['ari'])} & "
            f"{format_percent(row['gender_accuracy_if_2'])} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def plot_cmeans_2d(x: np.ndarray, y: np.ndarray, labels: np.ndarray, centers_scaled: np.ndarray, scaler: StandardScaler, path: Path, title: str) -> None:
    centers = scaler.inverse_transform(centers_scaled)
    plt.figure(figsize=(7, 5))
    plt.scatter(x[:, 0], x[:, 1], c=labels, cmap="Set2", s=34, alpha=0.82, edgecolor="white", linewidth=0.3)
    plt.scatter(centers[:, 0], centers[:, 1], c="black", marker="x", s=130, linewidths=2.2, label="centers")
    for label, marker in [("F", "o"), ("M", "^")]:
        mask = y == label
        plt.scatter([], [], marker=marker, color="gray", label=f"true {label}")
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title(title)
    plt.legend()
    plt.grid(alpha=0.22)
    plt.tight_layout()
    plt.savefig(path)
    plt.close()


def plot_hierarchical_2d(x: np.ndarray, y: np.ndarray, labels: np.ndarray, path: Path, title: str) -> None:
    plt.figure(figsize=(7, 5))
    plt.scatter(x[:, 0], x[:, 1], c=labels, cmap="Set2", s=34, alpha=0.82, edgecolor="white", linewidth=0.3)
    for label, marker in [("F", "o"), ("M", "^")]:
        plt.scatter([], [], marker=marker, color="gray", label=f"true {label}")
    plt.xlabel("Height")
    plt.ylabel("Weight")
    plt.title(title)
    plt.legend()
    plt.grid(alpha=0.22)
    plt.tight_layout()
    plt.savefig(path)
    plt.close()


def plot_k_metrics(rows: List[Dict[str, object]], dataset_name: str, path: Path) -> None:
    subset = [row for row in rows if row["dataset"] == dataset_name and row["method"] == "C-means"]
    subset = sorted(subset, key=lambda row: int(row["n_clusters"]))
    ks = [int(row["n_clusters"]) for row in subset]
    inertias = [float(row["inertia"]) for row in subset]
    silhouettes = [float(row["silhouette"]) for row in subset]
    fig, ax1 = plt.subplots(figsize=(7, 4.5))
    ax1.plot(ks, inertias, marker="o", color="#1b9e77", label="SSE")
    ax1.set_xlabel("Number of clusters C")
    ax1.set_ylabel("SSE", color="#1b9e77")
    ax1.tick_params(axis="y", labelcolor="#1b9e77")
    ax2 = ax1.twinx()
    ax2.plot(ks, silhouettes, marker="s", color="#d95f02", label="Silhouette")
    ax2.set_ylabel("Silhouette", color="#d95f02")
    ax2.tick_params(axis="y", labelcolor="#d95f02")
    plt.title(f"C-means Metrics on {dataset_name}")
    fig.tight_layout()
    plt.savefig(path)
    plt.close()


def plot_init_variation(rows: List[Dict[str, object]]) -> None:
    subset = [row for row in rows if row["method"] == "C-means different init"]
    seeds = [int(row["seed"]) for row in subset]
    inertias = [float(row["inertia"]) for row in subset]
    accuracies = [float(row["gender_accuracy_if_2"]) for row in subset]
    fig, ax1 = plt.subplots(figsize=(7, 4.5))
    ax1.bar(seeds, inertias, color="#80cdc1", alpha=0.85)
    ax1.set_xlabel("Initialization seed")
    ax1.set_ylabel("SSE")
    ax2 = ax1.twinx()
    ax2.plot(seeds, accuracies, color="#b2182b", marker="o")
    ax2.set_ylabel("Gender matching accuracy")
    plt.title("Effect of Different Initial Centers")
    fig.tight_layout()
    plt.savefig(FIG_DIR / "init_variation.pdf")
    plt.close()


def plot_dendrogram(x_scaled: np.ndarray, path: Path) -> None:
    linked = linkage(x_scaled, method="ward")
    plt.figure(figsize=(9, 4.8))
    dendrogram(linked, truncate_mode="lastp", p=20, leaf_rotation=35, leaf_font_size=8)
    plt.title("Hierarchical Clustering Dendrogram")
    plt.xlabel("Cluster")
    plt.ylabel("Distance")
    plt.tight_layout()
    plt.savefig(path)
    plt.close()


def main() -> None:
    ensure_dirs()
    data = load_data()
    all_rows: List[Dict[str, object]] = []
    scaled_data = {}
    scalers = {}

    for dataset_name in ["train", "combined"]:
        x, y = data[dataset_name]
        scaler = StandardScaler()
        x_scaled = scaler.fit_transform(x)
        scaled_data[dataset_name] = (x_scaled, y)
        scalers[dataset_name] = scaler
        all_rows.extend(evaluate_k_range(x_scaled, y, dataset_name))
        all_rows.append(evaluate_hierarchical(x_scaled, y, dataset_name))

    train_x, train_y = data["train"]
    train_scaled, _ = scaled_data["train"]
    all_rows.extend(evaluate_initial_seeds(train_scaled, train_y))

    labels, centers, inertia, _ = c_means(train_scaled, 2, seed=2028)
    plot_cmeans_2d(train_x, train_y, labels, centers, scalers["train"], FIG_DIR / "cmeans_train_2clusters.pdf", "C-means Clustering on Training Data")
    plot_k_metrics(all_rows, "train", FIG_DIR / "cmeans_k_metrics_train.pdf")
    plot_init_variation(all_rows)
    plot_dendrogram(train_scaled, FIG_DIR / "hierarchical_dendrogram_train.pdf")
    train_hier_labels = AgglomerativeClustering(n_clusters=2, linkage="ward").fit_predict(train_scaled)
    plot_hierarchical_2d(train_x, train_y, train_hier_labels, FIG_DIR / "hierarchical_train_2clusters.pdf", "Ward Hierarchical Clustering on Training Data")

    combined_x, combined_y = data["combined"]
    combined_scaled, _ = scaled_data["combined"]
    combined_labels, combined_centers, _, _ = c_means(combined_scaled, 2, seed=2028)
    plot_cmeans_2d(combined_x, combined_y, combined_labels, combined_centers, scalers["combined"], FIG_DIR / "cmeans_combined_2clusters.pdf", "C-means Clustering on Training + Test2")
    plot_k_metrics(all_rows, "combined", FIG_DIR / "cmeans_k_metrics_combined.pdf")
    combined_hier_labels = AgglomerativeClustering(n_clusters=2, linkage="ward").fit_predict(combined_scaled)
    plot_hierarchical_2d(combined_x, combined_y, combined_hier_labels, FIG_DIR / "hierarchical_combined_2clusters.pdf", "Ward Hierarchical Clustering on Training + Test2")

    write_csv(OUT_DIR / "metrics.csv", all_rows)
    write_summary_table(OUT_DIR / "summary_table.tex", all_rows)
    best_seed_rows = [row for row in all_rows if row["method"] == "C-means different init"]
    with (OUT_DIR / "cluster_notes.txt").open("w", encoding="utf-8") as f:
        f.write(f"train_samples: {len(train_x)}\n")
        f.write(f"combined_samples: {len(combined_x)}\n")
        f.write(f"init_sse_min: {min(float(row['inertia']) for row in best_seed_rows):.6f}\n")
        f.write(f"init_sse_max: {max(float(row['inertia']) for row in best_seed_rows):.6f}\n")
        f.write(f"init_accuracy_min: {min(float(row['gender_accuracy_if_2']) for row in best_seed_rows):.6f}\n")
        f.write(f"init_accuracy_max: {max(float(row['gender_accuracy_if_2']) for row in best_seed_rows):.6f}\n")

    print(f"Wrote metrics to {OUT_DIR / 'metrics.csv'}")
    print(f"Wrote figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
