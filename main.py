import os
import sys
import time

# Import modules from src
# Ensure __init__.py exists in src/ folder
try:
    from src import data_generator as generator
    from src import predict as predictor
    from src import train_model as trainer
except ImportError:
    # Handle import if running directly or file naming issues
    # Renaming imports for cleaner access if Python complains about numbers in filenames
    import importlib

    generator = importlib.import_module("src.01_data_generator")
    trainer = importlib.import_module("src.02_train_model")
    predictor = importlib.import_module("src.03_predict")


def clear_screen():
    os.system("cls" if os.name == "nt" else "clear")


def print_header():
    print("==========================================")
    print("   AI ANTENNA DESIGNER - CLI DASHBOARD    ")
    print("==========================================")
    print(" 1. [GENERATE] Run CST Automation & Collect Data")
    print(" 2. [TRAIN]    Train AI Model on CSV Data")
    print(" 3. [PREDICT]  Synthesize Antenna for Target Freq")
    print(" 4. [EXIT]     Quit Application")
    print("------------------------------------------")


def main():
    while True:
        clear_screen()
        print_header()

        choice = input("Select an option (1-4): ").strip()

        if choice == "1":
            print("\n--- DATA GENERATION MODE ---")
            try:
                n_str = input("How many samples to generate? (default 10): ")
                n = int(n_str) if n_str.isdigit() else 10

                print(
                    f"\n[INFO] Starting generation of {n} samples with VERBOSE LOGGING."
                )
                print(
                    "[INFO] Please ensure CST Studio is OPEN with your project loaded.\n"
                )

                confirm = input("Press ENTER to start (or 'q' to cancel)...")
                if confirm.lower() != "q":
                    generator.run_generator(num_samples=n, verbose=True)
                    input("\n[DONE] Press Enter to return to menu...")
            except Exception as e:
                print(f"\n[ERROR] Automation crashed: {e}")
                input("Press Enter to continue...")

        elif choice == "2":
            print("\n--- MODEL TRAINING MODE ---")
            try:
                success = trainer.train_model(verbose=True)
                if success:
                    print("\n[SUCCESS] Model is ready for predictions.")
                else:
                    print("\n[FAIL] Training failed.")
                input("\nPress Enter to return to menu...")
            except Exception as e:
                print(f"\n[ERROR] Training crashed: {e}")
                input("Press Enter to continue...")

        elif choice == "3":
            print("\n--- PREDICTION MODE ---")
            try:
                freq_str = input("Enter Target Resonance Frequency (in GHz): ")
                try:
                    freq = float(freq_str)
                    predictor.predict_design(freq, verbose=True)
                except ValueError:
                    print("[ERROR] Invalid number format.")

                input("Press Enter to return to menu...")
            except Exception as e:
                print(f"\n[ERROR] Prediction crashed: {e}")
                input("Press Enter to continue...")

        elif choice == "4":
            print("\nExiting... Good luck with your project!")
            sys.exit()

        else:
            input("\nInvalid option. Press Enter to try again...")


if __name__ == "__main__":
    main()
