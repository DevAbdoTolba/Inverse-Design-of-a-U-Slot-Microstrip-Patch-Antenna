import win32com.client
import sys

# Connect to the CST Project you just created manually
try:
    cst = win32com.client.GetActiveObject("CSTStudio.Application")
    mws = cst.Active3D()
except Exception as e:
    print("ERROR: Could not find an open CST project.")
    print("Please complete your Steps 1-13 first and keep CST open.")
    sys.exit()

print("Connected to active CST project.")

# --- 1. DEFINE PARAMETERS ---
# We define the variables that the AI will optimize later
params = {
    "W": 30.0,       # Patch Width
    "L": 28.0,       # Patch Length
    "Ls": 12.0,      # U-Slot Base Length
    "Ws": 2.0,       # U-Slot Thickness
    "La": 18.0,      # U-Slot Arm Length
    "h": 1.6,        # Substrate Height
    "W_sub": 60.0,   # Substrate Width
    "L_sub": 60.0    # Substrate Length
}

for name, val in params.items():
    mws.StoreParameter(name, val)

print("Parameters defined.")

# --- 2. GENERATE VBA MACRO FOR GEOMETRY ---
# We use the same VBA injection method because it's the most robust way to draw
# without getting COM errors.

vba_macro = """
' 1. CREATE SUBSTRATE
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

' 2. CREATE GROUND PLANE
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

' 3. CREATE RADIATING PATCH
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

' 4. CREATE U-SLOT (Using Vacuum Bricks)
' Base of U
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

' Left Arm
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

' Right Arm
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

' 5. BOOLEAN SUBTRACT
' Combine the slot parts first
Solid.Add "component1:Slot_Base", "component1:Slot_Arm_L"
Solid.Add "component1:Slot_Base", "component1:Slot_Arm_R"

' Cut them from the patch
Solid.Subtract "component1:Patch", "component1:Slot_Base"

' 6. DISCRETE PORT (FEED)
' We place a standard probe feed.
' Note: In a real optimized design, the feed position might also need to be a parameter (xf, yf).
' For now, we fix it at y = -8mm to be simple.
With DiscretePort 
    .Reset 
    .PortNumber "1" 
    .Type "SParameter" 
    .Label "" 
    .Impedance "50.0" 
    .SetP1 "0", "-8", "0" 
    .SetP2 "0", "-8", "h" 
    .Create 
End With

' Update the view
Rebuild
"""

# Send the VBA code to CST
mws.AddToHistory("Build Geometry", vba_macro)

print("Geometry built successfully! You can now check the 3D model in CST.")
