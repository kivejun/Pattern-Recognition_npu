# MNIST Dataset

Course source: http://yann.lecun.com/exdb/mnist/

This folder stores the standard IDX gzip files. Run:

```powershell
python scripts/prepare_datasets.py --download-small
```

The script downloads the four standard files and writes `manifest.csv`.
The standard split is 60,000 training images and 10,000 test images. If a validation
set is needed, use the last 10,000 training samples as validation and the first
50,000 as training.
