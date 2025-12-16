import win32com.client
import os
import time

# --- CONFIGURATION ---
# Make sure this matches exactly where you saved your file
PROJECT_PATH = os.path.abspath(os.path.join("cst_design", "patch_antenna.cst"))

print(f"Target Project: {PROJECT_PATH}")

if not os.path.exists(PROJECT_PATH):
    print("ERROR: File not found!")
    print("Please save your manually setup project to the path above first.")
    exit()

print("Launching new CST instance from Python...")

try:
    # 1. Start CST from Python (This ensures we have permission)
    cst = win32com.client.Dispatch("CSTStudio.Application")
    
    # 2. Open the file you saved
    print(f"Opening {PROJECT_PATH}...")
    cst.OpenFile(PROJECT_PATH)
    
    # Allow time for the file to load completely
    time.sleep(5) 
    
    # 3. Get the 3D Environment
    mws = cst.Active3D()
    print("SUCCESS: Connected to project!")

    # --- GEOMETRY BUILDER (Same as before) ---
    print("Defining Parameters...")
    params = {
        "W": 30.0, "L": 28.0, 
        "Ls": 12.0, "Ws": 2.0, "La": 18.0, 
        "h": 1.6, "W_sub": 60.0, "L_sub": 60.0
    }
    
    for name, val in params.items():
        mws.StoreParameter(name, val)

    print("Building U-Slot Geometry...")
    vba_macro = """
    ' 1. SUBSTRATE
    With Brick
        .Reset 
        .Name "Substrate" 
        .Component "component1" 
        .Material "FR-4 (lossy)" 
        .Xrange "-W_sub/2", "W_sub/2" 
        .Yrange "-L_sub/2", "L_sub/2" 
        .Zrange "0", "h" 
        .Create
    End With

    ' 2. GROUND
    With Brick
        .Reset 
        .Name "Ground" 
        .Component "component1" 
        .Material "PEC" 
        .Xrange "-W_sub/2", "W_sub/2" 
        .Yrange "-L_sub/2", "L_sub/2" 
        .Zrange "0", "0" 
        .Create
    End With

    ' 3. PATCH
    With Brick
        .Reset 
        .Name "Patch" 
        .Component "component1" 
        .Material "PEC" 
        .Xrange "-W/2", "W/2" 
        .Yrange "-L/2", "L/2" 
        .Zrange "h", "h" 
        .Create
    End With

    ' 4. SLOT CUTOUTS
    With Brick
        .Reset 
        .Name "Slot_Base" 
        .Component "component1" 
        .Material "Vacuum" 
        .Xrange "-Ls/2", "Ls/2" 
        .Yrange "-Ws/2", "Ws/2" 
        .Zrange "h", "h" 
        .Create
    End With

    With Brick
        .Reset 
        .Name "Slot_Arm_L" 
        .Component "component1" 
        .Material "Vacuum" 
        .Xrange "-Ls/2", "-Ls/2 + Ws" 
        .Yrange "0", "La" 
        .Zrange "h", "h" 
        .Create
    End With

    With Brick
        .Reset 
        .Name "Slot_Arm_R" 
        .Component "component1" 
        .Material "Vacuum" 
        .Xrange "Ls/2 - Ws", "Ls/2" 
        .Yrange "0", "La" 
        .Zrange "h", "h" 
        .Create
    End With

    ' 5. BOOLEAN
    Solid.Add "component1:Slot_Base", "component1:Slot_Arm_L"
    Solid.Add "component1:Slot_Base", "component1:Slot_Arm_R"
    Solid.Subtract "component1:Patch", "component1:Slot_Base"

    ' 6. PORT
    With DiscretePort 
        .Reset 
        .PortNumber "1" 
        .Type "SParameter" 
        .Impedance "50.0" 
        .SetP1 "0", "-8", "0" 
        .SetP2 "0", "-8", "h" 
        .Create 
    End With
    
    Rebuild
    """
    
    mws.AddToHistory("Build Geometry", vba_macro)
    
    print("Geometry built! Saving...")
    mws.Save()
    print("Done. You can now use the Data Generator.")

except Exception as e:
    print(f"CRITICAL ERROR: {e}")
