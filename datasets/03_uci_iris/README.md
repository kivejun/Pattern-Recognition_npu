# UCI Iris Dataset

Official source: https://archive.ics.uci.edu/dataset/53/iris

This folder is prepared for the UCI dataset requirement. Run:

```powershell
python scripts/prepare_datasets.py --download-small
```

The script downloads `iris.data` and writes `train.csv`, `val.csv`, and `test.csv`.
Split ratio: 60% / 20% / 20%.
