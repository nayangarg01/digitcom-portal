import sys
import os

# Add BillingEngine_OOP to python search path to import its modules
script_dir = os.path.dirname(os.path.abspath(__file__))
oop_dir = os.path.abspath(os.path.join(script_dir, "..", "..", "BillingEngine_OOP"))
sys.path.append(oop_dir)

# Import the OOP billing generator's main entry point
import oop_billing_generator

if __name__ == "__main__":
    oop_billing_generator.main()
