import sys
import os
import clr
import System
from System import Int32, Boolean, String, Double

# SETUP
POSSIBLE_PATHS = [
    os.path.join(os.getenv('LOCALAPPDATA'), 'DWSIM'),
    r"C:\Program Files\DWSIM",
    r"C:\Program Files (x86)\DWSIM"
]

DWSIM_PATH = None
for path in POSSIBLE_PATHS:
    if os.path.exists(os.path.join(path, "DWSIM.Automation.dll")):
        DWSIM_PATH = path
        break

if not DWSIM_PATH:
    sys.exit(1)

sys.path.append(DWSIM_PATH)
clr.AddReference("DWSIM.Automation")
clr.AddReference("DWSIM.Interfaces")
clr.AddReference("DWSIM.UnitOperations")

from DWSIM.Automation import Automation3
from DWSIM.Interfaces.Enums.GraphicObjects import ObjectType

interf = Automation3()
sim = interf.CreateFlowsheet()
pfr = sim.AddObject(ObjectType.RCT_PFR, Int32(100), Int32(100), String("PFR-Inspect"))

print("Listing PFR Properties via Reflection:")
pfr_type = pfr.GetType()
props = pfr_type.GetProperties()

found = False
for p in props:
    if "Reaction" in p.Name:
        print(f" - {p.Name} (Type: {p.PropertyType})")
        found = True

if not found:
    print("No properties with 'Reaction' in name found.")
