"""Experiment 4 main task: Eigenfaces with K-L/PCA face recognition."""

from __future__ import annotations

import csv
from pathlib import Path
from typing import Dict, List, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from sklearn.datasets import fetch_olivetti_faces
from sklearn.decomposition import PCA
from sklearn.metrics import accuracy_score, confusion_matrix
from sklearn.neighbors import KNeighborsClassifier, NearestCentroid
from sklearn.pipeline import make_pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.svm import SVC


ROOT = Path(__file__).resolve().parents[1]
DATA_HOME = ROOT / "datasets" / "07_olivetti_faces"
OUT_DIR = ROOT / "results" / "experiment4_eigenfaces"
FIG_DIR = ROOT / "reports" / "figures" / "experiment4_faces"

IMAGE_SHAPE = (64, 64)
TRAIN_PER_PERSON = 6
N_COMPONENTS = 120
COMPONENT_GRID = [10, 20, 40, 60, 80, 120, 160]
RANDOM_STATE = 2026


def ensure_dirs() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    FIG_DIR.mkdir(parents=True, exist_ok=True)


def load_faces() -> tuple[np.ndarray, np.ndarray, np.ndarray]:
    # Olivetti Faces 与 ORL/AT&T 数据集形式接近：40 人，每人 10 张，64x64 灰度图。
    data = fetch_olivetti_faces(data_home=DATA_HOME, shuffle=False)
    return data.data.astype(np.float32), data.images.astype(np.float32), data.target.astype(int)


def split_by_person(x: np.ndarray, images: np.ndarray, y: np.ndarray) -> Dict[str, np.ndarray]:
    train_indices: List[int] = []
    test_indices: List[int] = []
    for person in sorted(np.unique(y)):
        indices = np.flatnonzero(y == person)
        train_indices.extend(indices[:TRAIN_PER_PERSON])
        test_indices.extend(indices[TRAIN_PER_PERSON:])
    train_indices_np = np.array(train_indices)
    test_indices_np = np.array(test_indices)
    return {
        "x_train": x[train_indices_np],
        "y_train": y[train_indices_np],
        "img_train": images[train_indices_np],
        "x_test": x[test_indices_np],
        "y_test": y[test_indices_np],
        "img_test": images[test_indices_np],
    }


def fit_eval_models(
    x_train: np.ndarray,
    y_train: np.ndarray,
    x_test: np.ndarray,
    y_test: np.ndarray,
    n_components: int,
) -> List[Dict[str, object]]:
    n_components = min(n_components, x_train.shape[0], x_train.shape[1])
    models = {
        "Eigenfaces+NearestCentroid": make_pipeline(
            StandardScaler(),
            PCA(n_components=n_components, whiten=True, random_state=RANDOM_STATE),
            NearestCentroid(),
        ),
        "Eigenfaces+1NN": make_pipeline(
            StandardScaler(),
            PCA(n_components=n_components, whiten=True, random_state=RANDOM_STATE),
            KNeighborsClassifier(n_neighbors=1),
        ),
        "Eigenfaces+SVM": make_pipeline(
            StandardScaler(),
            PCA(n_components=n_components, whiten=True, random_state=RANDOM_STATE),
            SVC(kernel="linear", C=1.0, random_state=RANDOM_STATE),
        ),
    }
    rows: List[Dict[str, object]] = []
    for name, model in models.items():
        model.fit(x_train, y_train)
        pred_train = model.predict(x_train)
        pred_test = model.predict(x_test)
        rows.append(
            {
                "method": name,
                "n_components": n_components,
                "train_accuracy": accuracy_score(y_train, pred_train),
                "test_accuracy": accuracy_score(y_test, pred_test),
            }
        )
    return rows


def fit_main_model(
    x_train: np.ndarray,
    y_train: np.ndarray,
    x_test: np.ndarray,
) -> tuple[StandardScaler, PCA, SVC, np.ndarray, np.ndarray, np.ndarray]:
    scaler = StandardScaler()
    x_train_scaled = scaler.fit_transform(x_train)
    x_test_scaled = scaler.transform(x_test)
    pca = PCA(n_components=N_COMPONENTS, whiten=True, random_state=RANDOM_STATE)
    z_train = pca.fit_transform(x_train_scaled)
    z_test = pca.transform(x_test_scaled)
    clf = SVC(kernel="linear", C=1.0, random_state=RANDOM_STATE)
    clf.fit(z_train, y_train)
    y_pred = clf.predict(z_test)
    return scaler, pca, clf, z_train, z_test, y_pred


def write_csv(path: Path, rows: Sequence[Dict[str, object]]) -> None:
    with path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)


def format_percent(value: float) -> str:
    return f"{value * 100:.2f}\\%"


def write_summary_table(path: Path, rows: Sequence[Dict[str, object]]) -> None:
    lines = [
        "\\begin{tabular}{lccc}",
        "\\toprule",
        "方法 & 主成分数 & 训练准确率 & 测试准确率 \\\\",
        "\\midrule",
    ]
    selected = [row for row in rows if int(row["n_components"]) == N_COMPONENTS]
    for row in selected:
        lines.append(
            f"{row['method']} & {int(row['n_components'])} & "
            f"{format_percent(float(row['train_accuracy']))} & "
            f"{format_percent(float(row['test_accuracy']))} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def write_component_table(
    path: Path,
    rows: Sequence[Dict[str, object]],
    cumulative_variance: np.ndarray,
) -> None:
    by_n: Dict[int, Dict[str, Dict[str, object]]] = {}
    for row in rows:
        by_n.setdefault(int(row["n_components"]), {})[str(row["method"])] = row
    lines = [
        "\\begin{tabular}{crrrr}",
        "\\toprule",
        "主成分数 & 累计方差 & 最近质心 & 1NN & SVM \\\\",
        "\\midrule",
    ]
    for n_components in sorted(by_n):
        nearest = by_n[n_components]["Eigenfaces+NearestCentroid"]
        knn = by_n[n_components]["Eigenfaces+1NN"]
        svm = by_n[n_components]["Eigenfaces+SVM"]
        lines.append(
            f"{n_components} & {format_percent(float(cumulative_variance[n_components - 1]))} & "
            f"{format_percent(float(nearest['test_accuracy']))} & "
            f"{format_percent(float(knn['test_accuracy']))} & "
            f"{format_percent(float(svm['test_accuracy']))} \\\\"
        )
    lines += ["\\bottomrule", "\\end{tabular}", ""]
    path.write_text("\n".join(lines), encoding="utf-8")


def write_dataset_table(path: Path) -> None:
    lines = [
        "\\begin{tabular}{lccc}",
        "\\toprule",
        "数据集 & 人数 & 每人图像数 & 图像大小 \\\\",
        "\\midrule",
        "Olivetti Faces & 40 & 10 & $64\\times64$ \\\\",
        "\\bottomrule",
        "\\end{tabular}",
        "",
    ]
    path.write_text("\n".join(lines), encoding="utf-8")


def plot_face_samples(images: np.ndarray, labels: np.ndarray) -> None:
    selected = []
    for person in range(16):
        selected.append(np.flatnonzero(labels == person)[0])
    fig, axes = plt.subplots(4, 4, figsize=(6.4, 6.4))
    for ax, idx in zip(axes.ravel(), selected):
        ax.imshow(images[idx], cmap="gray")
        ax.set_title(f"Person {labels[idx]}", fontsize=8)
        ax.axis("off")
    plt.suptitle("Olivetti Face Samples", fontsize=12)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "face_samples.pdf")
    plt.close()


def plot_average_face(x_train: np.ndarray) -> None:
    mean_face = x_train.mean(axis=0).reshape(IMAGE_SHAPE)
    plt.figure(figsize=(3.4, 3.4))
    plt.imshow(mean_face, cmap="gray")
    plt.title("Mean Face")
    plt.axis("off")
    plt.tight_layout()
    plt.savefig(FIG_DIR / "mean_face.pdf")
    plt.close()


def plot_eigenfaces(pca: PCA) -> None:
    fig, axes = plt.subplots(3, 5, figsize=(8, 5.2))
    for idx, ax in enumerate(axes.ravel()):
        comp = pca.components_[idx].reshape(IMAGE_SHAPE)
        ax.imshow(comp, cmap="gray")
        ax.set_title(f"EF{idx + 1}", fontsize=8)
        ax.axis("off")
    plt.suptitle("First Eigenfaces", fontsize=12)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "eigenfaces.pdf")
    plt.close()


def plot_explained_variance(cumulative_variance: np.ndarray) -> None:
    xs = np.arange(1, len(cumulative_variance) + 1)
    fig, ax = plt.subplots(figsize=(7.4, 4.5))
    ax.plot(xs, cumulative_variance, linewidth=2)
    ax.scatter(COMPONENT_GRID, cumulative_variance[np.array(COMPONENT_GRID) - 1], color="#d62728", zorder=3)
    ax.axhline(0.9, color="#555555", linestyle="--", linewidth=1)
    ax.set_xlabel("Number of eigenfaces")
    ax.set_ylabel("Cumulative explained variance")
    ax.set_title("Eigenfaces Cumulative Explained Variance")
    ax.set_ylim(0, 1.02)
    ax.grid(alpha=0.25)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "explained_variance.pdf")
    plt.close()


def plot_accuracy_curve(rows: Sequence[Dict[str, object]]) -> None:
    fig, ax = plt.subplots(figsize=(7.4, 4.6))
    for method in sorted(set(str(row["method"]) for row in rows)):
        subset = sorted([row for row in rows if row["method"] == method], key=lambda row: int(row["n_components"]))
        xs = [int(row["n_components"]) for row in subset]
        ys = [float(row["test_accuracy"]) for row in subset]
        ax.plot(xs, ys, marker="o", linewidth=1.8, label=method.replace("Eigenfaces+", ""))
    ax.set_xlabel("Number of eigenfaces")
    ax.set_ylabel("Test accuracy")
    ax.set_title("Face Recognition Accuracy with Eigenfaces")
    ax.set_ylim(0, 1.02)
    ax.grid(alpha=0.25)
    ax.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "accuracy_vs_components.pdf")
    plt.close()


def plot_projection(z_test: np.ndarray, y_true: np.ndarray, y_pred: np.ndarray) -> None:
    correct = y_true == y_pred
    fig, ax = plt.subplots(figsize=(7, 5.2))
    ax.scatter(z_test[correct, 0], z_test[correct, 1], s=22, alpha=0.8, label="Correct", color="#2ca02c")
    ax.scatter(z_test[~correct, 0], z_test[~correct, 1], s=28, alpha=0.9, label="Wrong", color="#d62728", marker="x")
    ax.set_xlabel("Eigenface 1 coefficient")
    ax.set_ylabel("Eigenface 2 coefficient")
    ax.set_title("Test Faces in the First Two Eigenface Dimensions")
    ax.grid(alpha=0.25)
    ax.legend()
    plt.tight_layout()
    plt.savefig(FIG_DIR / "projection_correct_wrong.pdf")
    plt.close()


def plot_confusion(y_true: np.ndarray, y_pred: np.ndarray) -> None:
    classes = sorted(np.unique(y_true))
    cm = confusion_matrix(y_true, y_pred, labels=classes)
    cm_norm = cm / np.maximum(cm.sum(axis=1, keepdims=True), 1)
    fig, ax = plt.subplots(figsize=(8.2, 7.2))
    im = ax.imshow(cm_norm, cmap="Blues", vmin=0, vmax=1)
    ax.set_xticks(np.arange(len(classes)), classes, fontsize=6)
    ax.set_yticks(np.arange(len(classes)), classes, fontsize=6)
    ax.set_xlabel("Predicted person")
    ax.set_ylabel("True person")
    ax.set_title("Eigenfaces+SVM Normalized Confusion Matrix")
    fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "confusion_matrix.pdf")
    plt.close()


def plot_prediction_examples(images: np.ndarray, y_true: np.ndarray, y_pred: np.ndarray) -> None:
    correct_indices = np.flatnonzero(y_true == y_pred)[:8]
    wrong_indices = np.flatnonzero(y_true != y_pred)[:8]
    selected = list(correct_indices) + list(wrong_indices)
    fig, axes = plt.subplots(4, 4, figsize=(7.4, 7.8))
    for ax, idx in zip(axes.ravel(), selected):
        ax.imshow(images[int(idx)], cmap="gray")
        ok = y_true[int(idx)] == y_pred[int(idx)]
        color = "#2ca02c" if ok else "#d62728"
        ax.set_title(f"T:{y_true[int(idx)]} P:{y_pred[int(idx)]}", fontsize=8, color=color)
        ax.axis("off")
    for ax in axes.ravel()[len(selected) :]:
        ax.axis("off")
    plt.suptitle("Eigenfaces+SVM Prediction Examples", fontsize=12)
    plt.tight_layout()
    plt.savefig(FIG_DIR / "prediction_examples.pdf")
    plt.close()


def main() -> None:
    ensure_dirs()
    x, images, y = load_faces()
    split = split_by_person(x, images, y)
    x_train = split["x_train"]
    y_train = split["y_train"]
    x_test = split["x_test"]
    y_test = split["y_test"]

    scaler_for_variance = StandardScaler()
    x_train_scaled = scaler_for_variance.fit_transform(x_train)
    pca_full = PCA(n_components=max(COMPONENT_GRID), whiten=True, random_state=RANDOM_STATE)
    pca_full.fit(x_train_scaled)
    cumulative_variance = np.cumsum(pca_full.explained_variance_ratio_)

    rows: List[Dict[str, object]] = []
    for n_components in COMPONENT_GRID:
        rows.extend(fit_eval_models(x_train, y_train, x_test, y_test, n_components))

    scaler, pca, clf, z_train, z_test, y_pred = fit_main_model(x_train, y_train, x_test)

    write_csv(OUT_DIR / "metrics.csv", rows)
    write_dataset_table(OUT_DIR / "dataset_table.tex")
    write_summary_table(OUT_DIR / "summary_table.tex", rows)
    write_component_table(OUT_DIR / "component_table.tex", rows, cumulative_variance)

    plot_face_samples(images, y)
    plot_average_face(x_train)
    plot_eigenfaces(pca)
    plot_explained_variance(cumulative_variance)
    plot_accuracy_curve(rows)
    plot_projection(z_test, y_test, y_pred)
    plot_confusion(y_test, y_pred)
    plot_prediction_examples(split["img_test"], y_test, y_pred)

    with (OUT_DIR / "pca_params.txt").open("w", encoding="utf-8") as f:
        f.write("dataset: Olivetti Faces\n")
        f.write("num_persons: 40\n")
        f.write("images_per_person: 10\n")
        f.write(f"train_per_person: {TRAIN_PER_PERSON}\n")
        f.write(f"train_samples: {len(x_train)}\n")
        f.write(f"test_samples: {len(x_test)}\n")
        f.write(f"image_shape: {IMAGE_SHAPE[0]}x{IMAGE_SHAPE[1]}\n")
        f.write(f"n_components: {N_COMPONENTS}\n")
        f.write(f"explained_variance_ratio_sum: {pca.explained_variance_ratio_.sum():.6f}\n")
        f.write(f"main_model: Eigenfaces+SVM\n")
        f.write(f"test_accuracy: {accuracy_score(y_test, y_pred):.6f}\n")

    print(f"Wrote eigenface results to {OUT_DIR}")
    print(f"Wrote eigenface figures to {FIG_DIR}")


if __name__ == "__main__":
    main()
