import os
import random
import time
from datetime import datetime

import pandas as pd
import win32com.client

# Configuration
CSV_PATH = os.path.join("data", "antenna_data.csv")
PARAM_BOUNDS = {
    "W": (30.0, 50.0),
    "L": (25.0, 40.0),
    "Ws": (2.0, 8.0),
    "Ls": (10.0, 20.0),
}


def log(msg, verbose):
    if verbose:
        print(f"[{datetime.now().strftime('%H:%M:%S')}] [GEN] {msg}")


def run_generator(num_samples=10, verbose=True):
    log("Initializing Data Generator...", verbose)

    # Check directory
    if not os.path.exists("data"):
        os.makedirs("data")
        log("Created 'data' directory.", verbose)

    # Connect to CST
    try:
        log("Attempting to connect to CST Studio COM Interface...", verbose)
        cst = win32com.client.Dispatch("CSTStudio.Application")
        mws = cst.Active3D()
        if mws is None:
            raise Exception("CST Active3D Object is None.")
        log("Connected to Active CST Project.", verbose)
    except Exception as e:
        log(f"CRITICAL ERROR: Could not connect to CST. {e}", True)
        log("Ensure CST is open and a project is loaded.", True)
        return False

    data_log = []

    # Load existing data if available to append
    if os.path.exists(CSV_PATH):
        try:
            existing_df = pd.read_csv(CSV_PATH)
            data_log = existing_df.to_dict("records")
            log(f"Loaded {len(data_log)} existing samples from CSV.", verbose)
        except:
            pass

    log(f"Starting loop for {num_samples} new samples...", verbose)

    for i in range(num_samples):
        try:
            # 1. Random Parameters
            W = round(random.uniform(*PARAM_BOUNDS["W"]), 2)
            L = round(random.uniform(*PARAM_BOUNDS["L"]), 2)

            # Constraints
            max_Ls = W - 4.0
            max_Ws = (L / 2) - 2.0

            Ls = round(random.uniform(10.0, max_Ls), 2)
            Ws = round(random.uniform(2.0, max_Ws), 2)

            params = {"W": W, "L": L, "Ls": Ls, "Ws": Ws}

            log(f"Iter {i + 1}: Applying Params {params}", verbose)

            # 2. Update CST
            for key, val in params.items():
                mws.StoreParameter(key, val)

            # 3. Rebuild
            log(f"Iter {i + 1}: Rebuilding Geometry...", verbose)
            mws.RebuildOnParametricChange(False, False)

            # 4. Solve
            log(f"Iter {i + 1}: Starting Solver (This may take time)...", verbose)
            solver = mws.Solver
            solver.Start()

            # 5. Extract Results
            result_tree = mws.ResultTree
            s11_obj = result_tree.GetResultFromTreeItem(
                "1D Results\\S-Parameters\\S1,1", "3D:RunID:0"
            )

            if s11_obj:
                mags = s11_obj.GetResultValuesY()
                freqs = s11_obj.GetResultValuesX()

                min_s11 = min(mags)
                min_idx = mags.index(min_s11)
                res_freq = freqs[min_idx]

                row = params.copy()
                row["res_freq"] = res_freq
                row["s11_min"] = min_s11
                data_log.append(row)
                log(
                    f"Iter {i + 1}: SUCCESS -> Freq={res_freq:.2f}GHz, S11={min_s11:.2f}dB",
                    verbose,
                )
            else:
                log(f"Iter {i + 1}: FAILED -> No S11 results found.", verbose)

        except Exception as e:
            log(f"Iter {i + 1}: ERROR -> {str(e)}", verbose)

        # Periodic Save
        if i % 5 == 0:
            pd.DataFrame(data_log).to_csv(CSV_PATH, index=False)
            log("Progress saved to CSV.", verbose)

    # Final Save
    pd.DataFrame(data_log).to_csv(CSV_PATH, index=False)
    log("Data Generation Complete.", verbose)
    return True
