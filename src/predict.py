import os
import warnings

import joblib
import pandas as pd

warnings.filterwarnings("ignore")

MODEL_PATH = os.path.join("models", "antenna_model.pkl")


def predict_design(target_freq, verbose=True):
    if verbose:
        print(f"\n[AI] Loading Model from {MODEL_PATH}...")

    if not os.path.exists(MODEL_PATH):
        print("[ERROR] Model file not found! Train the model first.")
        return

    model = joblib.load(MODEL_PATH)

    input_data = pd.DataFrame([[target_freq]], columns=["res_freq"])

    if verbose:
        print(f"[AI] Predicting geometry for target: {target_freq} GHz...")

    prediction = model.predict(input_data)
    dims = prediction[0]

    print("\n" + "=" * 40)
    print(f"  AI SYNTHESIS RESULT: {target_freq} GHz")
    print("=" * 40)
    print(f"  Patch Width (W)  : {dims[0]:.3f} mm")
    print(f"  Patch Length (L) : {dims[1]:.3f} mm")
    print(f"  Slot Length (Ls) : {dims[2]:.3f} mm")
    print(f"  Slot Width (Ws)  : {dims[3]:.3f} mm")
    print("=" * 40 + "\n")
