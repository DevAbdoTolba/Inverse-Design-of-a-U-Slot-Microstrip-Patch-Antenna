import win32com.client
import sys

print("--- CST CONNECTION DIAGNOSTIC ---")

try:
    # Try the modern ID
    print("Attempting 'CSTStudio.Application'...")
    cst = win32com.client.GetActiveObject("CSTStudio.Application")
    print("SUCCESS: Connected via CSTStudio.Application")
    
    # Test access to Active3D
    mws = cst.Active3D()
    if mws:
        print("SUCCESS: Active3D object retrieved.")
    else:
        print("WARNING: Connected to App, but no Active3D project found.")

except Exception as e:
    print(f"FAILED: {e}")
    
    print("\nAttempting legacy ID 'CSTDesignEnvironment.Application'...")
    try:
        # Try the legacy ID (sometimes used in older/specific installations)
        cst = win32com.client.GetActiveObject("CSTDesignEnvironment.Application")
        print("SUCCESS: Connected via CSTDesignEnvironment.Application")
    except Exception as e2:
        print(f"FAILED: {e2}")

print("\n---------------------------------")
print("TROUBLESHOOTING:")
print("1. If you see 'Operation unavailable', run Python as ADMINISTRATOR.")
print("2. If you see 'Invalid class string', CST might not be registered in Windows Registry properly.")
