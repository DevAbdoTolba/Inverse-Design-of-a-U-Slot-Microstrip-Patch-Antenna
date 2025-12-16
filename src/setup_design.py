import os

import win32com.client

# --- CONFIGURATION ---
PROJECT_NAME = "patch_antenna.cst"
DESIGN_DIR = os.path.join(os.path.dirname(__file__), "..", "cst_design")
FULL_PATH = os.path.join(DESIGN_DIR, PROJECT_NAME)

# Default Dimensions (in mm) to initialize the file
# These will be overwritten by the AI later, but we need a valid starting shape.
INIT_PARAMS = {
    "W": 30.0,  # Patch Width
    "L": 28.0,  # Patch Length
    "Ls": 12.0,  # U-Slot Base Length
    "Ws": 2.0,  # U-Slot Thickness (width of the cut)
    "La": 18.0,  # U-Slot Arm Length (Added this as it's required for U-shape)
    "h": 1.6,  # Substrate Height
    "W_sub": 60.0,  # Substrate Width
    "L_sub": 60.0,  # Substrate Length
}


def create_cst_project():
    if not os.path.exists(DESIGN_DIR):
        os.makedirs(DESIGN_DIR)

    print("Launching CST Studio...")
    cst = win32com.client.Dispatch("CSTStudio.Application")
    cst.NewMWS()  # Create new Microwave Studio Project
    mws = cst.Active3D()

    print("Setting up Units and Solver...")
    # Set Units to mm, GHz, ns
    mws.AddToHistory(
        "define units",
        """
        With Units
            .Geometry "mm"
            .Frequency "GHz"
            .Time "ns"
        End With
    """,
    )

    # Set Frequency Range (1 to 5 GHz)
    mws.AddToHistory("define frequency range", 'Solver.FrequencyRange "1", "5"')

    print("Defining Parameters...")
    # Add parameters to the CST list
    for name, val in INIT_PARAMS.items():
        mws.StoreParameter(name, val)

    print("Building Geometry...")

    # 1. Create Substrate (FR4)
    # VBA command to create a brick
    vba_substrate = """
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
    """
    mws.AddToHistory("create substrate", vba_substrate)

    # 2. Create Ground Plane (PEC)
    vba_ground = """
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
    """
    mws.AddToHistory("create ground", vba_ground)

    # 3. Create Patch (PEC)
    vba_patch = """
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
    """
    mws.AddToHistory("create patch", vba_patch)

    # 4. Create U-Slot (This is the tricky part - Boolean Subtract)
    # We create a "SlotTool" object and subtract it from the Patch

    # Slot Base (Horizontal part of U)
    vba_slot_base = """
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
    """
    mws.AddToHistory("create slot base", vba_slot_base)

    # Slot Arm Left
    vba_slot_arm_L = """
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
    """
    mws.AddToHistory("create slot arm left", vba_slot_arm_L)

    # Slot Arm Right
    vba_slot_arm_R = """
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
    """
    mws.AddToHistory("create slot arm right", vba_slot_arm_R)

    # 5. Boolean Operations to form the U-Slot
    # Unite the 3 slot parts
    mws.AddToHistory(
        "unite slot parts",
        """
        Solid.Add "component1:Slot_Base", "component1:Slot_Arm_L"
        Solid.Add "component1:Slot_Base", "component1:Slot_Arm_R"
    """,
    )

    # Subtract the Slot structure from the Patch
    mws.AddToHistory(
        "cut slot from patch",
        'Solid.Subtract "component1:Patch", "component1:Slot_Base"',
    )

    print("Adding Feed Port...")
    # 6. Create Discrete Port (Probe Feed)
    # Placed at (0, -10) mm - typical for U-slot patches, usually needs optimization
    mws.AddToHistory(
        "define discrete port",
        """
        With DiscretePort
            .Reset
            .PortNumber "1"
            .Type "SParameter"
            .Label ""
            .Folder ""
            .Impedance "50.0"
            .VoltagePort "False"
            .CurrentPort "False"
            .Monitor "True"
            .Radius "0.0"
            .SetP1 "0", "-8", "0"
            .SetP2 "0", "-8", "h"
            .InvertDirection "False"
            .LocalCoordinates "False"
            .Wire ""
            .Position "central"
            .Create
        End With
    """,
    )

    print(f"Saving project to {FULL_PATH}...")
    mws.SaveAs(FULL_PATH, True)
    print("Design Created Successfully!")


if __name__ == "__main__":
    create_cst_project()
