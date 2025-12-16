import win32com.client
import os
import time

# --- CONFIGURATION ---
PROJECT_PATH = os.path.abspath(os.path.join("cst_design", "patch_antenna.cst"))
TEMP_MACRO_PATH = os.path.abspath("temp_build_geometry.bas")

print(f"Target Project: {PROJECT_PATH}")

if not os.path.exists(PROJECT_PATH):
    print("ERROR: File not found! Save your manual project first.")
    exit()

print("Launching CST (Force Mode)...")

try:
    # 1. Start CST
    cst = win32com.client.Dispatch("CSTStudio.Application")
    cst.OpenFile(PROJECT_PATH)
    time.sleep(5) 
    mws = cst.Active3D()
    print("SUCCESS: Connected.")

    # 2. Define Parameters directly
    print("Setting Parameters...")
    params = {
        "W": 30.0, "L": 28.0, 
        "Ls": 12.0, "Ws": 2.0, "La": 18.0, 
        "h": 1.6, "W_sub": 60.0, "L_sub": 60.0
    }
    for name, val in params.items():
        mws.StoreParameter(name, val)

    # 3. WRITE VBA TO A FILE (The Fix)
    print("Generating VBA Macro File...")
    
    vba_code = """
    ' CST SCRIPT TO BUILD GEOMETRY
    
    Sub Main ()
        
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
    
        ' 4. SLOT PARTS
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
    
        ' 5. BOOLEAN OPERATIONS
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
        
        ' Update View
        Rebuild
        
    End Sub
    """

    # Save to temp file
    with open(TEMP_MACRO_PATH, "w") as f:
        f.write(vba_code)

    # 4. EXECUTE FILE
    print(f"Executing Macro: {TEMP_MACRO_PATH}")
    mws.RunMacro(TEMP_MACRO_PATH)
    
    print("Geometry built successfully!")
    mws.Save()

except Exception as e:
    print(f"CRITICAL ERROR: {e}")
