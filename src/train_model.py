import os
from datetime import datetime

import joblib
import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.metrics import mean_absolute_error, r2_score
from sklearn.model_selection import train_test_split
from sklearn.multioutput import MultiOutputRegressor

DATA_PATH = os.path.join("data", "antenna_data.csv")
MODEL_PATH = os.path.join("models", "antenna_model.pkl")


def log(msg, verbose):
    if verbose:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] [TRAIN] {msg}")


def train_model(verbose=True):
    log("Checking data availability...", verbose)

    if not os.path.exists(DATA_PATH):
        log(f"ERROR: Dataset not found at {DATA_PATH}", True)
        return False

    if not os.path.exists("models"):
        os.makedirs("models")

    # Load Data
    df = pd.read_csv(DATA_PATH)
    log(f"Loaded {len(df)} samples.", verbose)

    # Filter bad antennas (S11 must be decent)
    initial_count = len(df)
    df = df[df["s11_min"] < -5]  # Relaxed constraint for testing
    log(
        f"Filtered poor antennas. {len(df)}/{initial_count} usable samples remain.",
        verbose,
    )

    if len(df) < 10:
        log("ERROR: Not enough data to train. Need at least 10 valid samples.", True)
        return False

    # Features & Targets
    X = df[["res_freq"]]
    y = df[["W", "L", "Ls", "Ws"]]

    log("Splitting Train/Test data...", verbose)
    X_train, X_test, y_train, y_test = train_test_split(
        X, y, test_size=0.2, random_state=42
    )

    log("Initializing Random Forest Regressor...", verbose)
    rf = RandomForestRegressor(n_estimators=100, random_state=42)
    model = MultiOutputRegressor(rf)

    log("Fitting model...", verbose)
    model.fit(X_train, y_train)

    log("Evaluating model...", verbose)
    y_pred = model.predict(X_test)

    mae = mean_absolute_error(y_test, y_pred)
    r2 = r2_score(y_test, y_pred)

    log(f"Evaluation Results:\n   - MAE: {mae:.4f} mm\n   - R2 Score: {r2:.4f}", True)

    joblib.dump(model, MODEL_PATH)
    log(f"Model saved successfully to {MODEL_PATH}", True)
    return True
