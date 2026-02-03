import sys
import os
import clr
import time
import System
from System import Int32, Boolean, String, Double
from System.Collections.Generic import Dictionary

# ==============================================================================
# SETUP
# ==============================================================================

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
    print("Error: DWSIM not found.")
    sys.exit(1)

print(f"Found DWSIM at: {DWSIM_PATH}")
sys.path.append(DWSIM_PATH)

try:
    clr.AddReference("DWSIM.Automation")
    clr.AddReference("DWSIM.Interfaces")
    clr.AddReference("DWSIM.GlobalSettings")
    clr.AddReference("DWSIM.SharedClasses")
    clr.AddReference("DWSIM.Thermodynamics")
    clr.AddReference("DWSIM.UnitOperations")
    
    from DWSIM.Automation import Automation3
    from DWSIM.Interfaces.Enums.GraphicObjects import ObjectType
except Exception as e:
    print(f"Error loading assemblies: {e}")
    sys.exit(1)

# ==============================================================================
# DEBUG SIMULATION
# ==============================================================================

def run_debug():
    print("Initializing Automation...")
    interf = Automation3()
    sim = interf.CreateFlowsheet()
    
    # 1. Setup Compounds
    print("Adding Compounds...")
    try:
        sim.AddCompound("Water")
        sim.AddCompound("Ethanol")
        print("Compounds added.")
    except Exception as e:
        print(f"Error adding compounds: {e}")
        return

    # 2. Setup Feed
    print("Adding Feed...")
    feed = sim.AddObject(ObjectType.MaterialStream, Int32(50), Int32(50), String("Feed"))
    feed.SetPropertyValue("Temperature", 350.0)
    feed.SetPropertyValue("Pressure", 101325.0)
    feed.SetPropertyValue("MolarFlow", 100.0)
    feed.SetPropertyValue("OverallMolarComposition", [1.0, 0.0]) # 100% Water
    
    print("Adding Product/Energy...")
    prod = sim.AddObject(ObjectType.MaterialStream, Int32(200), Int32(50), String("Product"))
    energy = sim.AddObject(ObjectType.EnergyStream, Int32(125), Int32(100), String("Energy"))

    # 3. Define Reaction
    print("Defining Reaction...")
    stoich = Dictionary[String, Double]()
    stoich.Add("Water", -1.0)
    stoich.Add("Ethanol", 1.0)
    
    direct_orders = Dictionary[String, Double]()
    direct_orders.Add("Water", 1.0)
    direct_orders.Add("Ethanol", 0.0)
    
    rev_orders = Dictionary[String, Double]()
    rev_orders.Add("Water", 0.0)
    rev_orders.Add("Ethanol", 0.0)
    
    # High kinetic factor K
    rxn = sim.CreateKineticReaction(
        String("Rxn-Test"), 
        String("Water->Ethanol"), 
        stoich, 
        direct_orders, 
        rev_orders, 
        String("Water"), 
        String("Mixture"), 
        String("Molar Concentration"), 
        String("mol"), 
        String("mol/[m3.s]"), 
        Double(50.0),   
        Double(0.0),   
        Double(0.0), 
        Double(0.0), 
        String(""), 
        String("")
    )
    
    # 4. Reaction Set
    print("Configuring Reaction Set...")
    rxn_sets = sim.ReactionSets
    rs = None
    
    if rxn_sets.Count > 0:
        for key in rxn_sets.Keys:
            rs = rxn_sets[key]
            print(f"Using existing set: {rs.Name} (ID: {rs.ID})")
            break
    else:
        rs = sim.CreateReactionSet(String("DebugSet"), String("Active Set"))
        print(f"Created set: {rs.Name} (ID: {rs.ID})")
        
    try:
        sim.AddReactionToSet(rxn.ID, rs.ID, Boolean(True), Int32(0))
        print("Reaction added to set successfully.")
    except Exception as e:
        print(f"ERROR adding reaction to set: {e}")
        return

    # 5. PFR Setup
    print("Adding PFR...")
    pfr = sim.AddObject(ObjectType.RCT_PFR, Int32(125), Int32(50), String("PFR-Debug"))
    sim.ConnectObjects(feed.GraphicObject, pfr.GraphicObject, -1, -1)
    sim.ConnectObjects(pfr.GraphicObject, prod.GraphicObject, -1, -1)
    sim.ConnectObjects(energy.GraphicObject, pfr.GraphicObject, -1, -1)
    
    # 6. PFR Settings via Reflection
    print("Configuring PFR Properties...")
    pfr_type = pfr.GetType()
    
    try:
        # Volume
        prop_vol = pfr_type.GetProperty("Volume")
        if prop_vol: 
            prop_vol.SetValue(pfr, Double(5.0), None)
            print("PFR Volume set.")
            
        # Mode
        prop_mode = pfr_type.GetProperty("ReactorOperationMode")
        if prop_mode:
            prop_mode.SetValue(pfr, Int32(1), None)
            print("PFR Mode set.")
            
    except Exception as e:
        print(f"Prop setup error: {e}")
    
    # 7. Assign Reaction Set (Tricky Part)
    print("Assigning Reaction Set...")
    valid_index = -1
    idx = 0
    target_set_name = rs.Name
    
    for kvp in sim.ReactionSets:
        if kvp.Value.Name == target_set_name:
            valid_index = idx
        idx += 1
    
    if valid_index >= 0:
        print(f"Target Set Index: {valid_index}")
        assigned = False
        
        # Method 1: Property "Reaction Set" (Int)
        try:
            prop = pfr_type.GetProperty("Reaction Set")
            if prop:
                prop.SetValue(pfr, Int32(valid_index), None)
                print("Assigned via 'Reaction Set' property.")
                assigned = True
        except: pass
        
        if not assigned:
            # Method 2: Property "ReactionSet" (No Space)
            try:
                prop = pfr_type.GetProperty("ReactionSet")
                if prop:
                    prop.SetValue(pfr, Int32(valid_index), None)
                    print("Assigned via 'ReactionSet' property.")
                    assigned = True
            except: pass
            
        if not assigned:
             # Method 3: Direct Attribute (PythonNet)
             try:
                 pfr.ReactionSet = valid_index
                 print("Assigned via python attribute 'ReactionSet'.")
                 assigned = True
             except: pass
             
        if not assigned:
             # Method 4: Direct Attribute with Space? No.
             print("FAILED to assign Reaction Set via all known methods.")
    else:
        print("Set Index not found.")

    # 8. Solve
    print("Solving...")
    interf.CalculateFlowsheet2(sim)
    
    # 9. Results
    print("Reading Results...")
    try:
        # Try getting flows directly
        # If 'OverallMolarComposition' returns string, use other props
        feed_flow = feed.GetPropertyValue("MolarFlow")
        prod_flow = prod.GetPropertyValue("MolarFlow")
        print(f"Feed Flow: {feed_flow}, Prod Flow: {prod_flow}")
        
        comp_obj = prod.GetPropertyValue("OverallMolarComposition")
        print(f"Prod Composition (Raw): {comp_obj}")
        
        # If it's a string, we can't do much unless we parse it, but let's assume if it worked in original it should work here if calc succeeded.
        # But if it is failing, let's look at conversion via Mass Balance if Flows changed?
        # A->B Isomers conserve moles. Water->Ethanol does NOT conserve moles.
        # 1 mol Water (18g) -> 1 mol Ethanol (46g). Mass Creation!
        # DWSIM might fail to solve or give weird results.
        # Check Mass Flow
        feed_mass = feed.GetPropertyValue("MassFlow")
        prod_mass = prod.GetPropertyValue("MassFlow")
        print(f"Feed Mass: {feed_mass}, Prod Mass: {prod_mass}")
        
        if hasattr(comp_obj, '__getitem__') and len(comp_obj) > 1:
             prod_eth = comp_obj[1]
             print(f"Prod Ethanol Fraction: {prod_eth}")
             if prod_eth > 0.01:
                 print("SUCCESS: Conversion observed.")
             else:
                 print("FAILURE: Zero conversion.")
        elif isinstance(comp_obj, str):
             print(f"Got string composition: '{comp_obj}'")
             # Maybe use 'Combined Material Stream' properties?
    except Exception as e:
        print(f"Result Read Error: {e}")

if __name__ == "__main__":
    run_debug()
