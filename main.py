from scripts import driver_list
from scripts.driver_list import list_intel_drivers

import json
from scripts.driver_list import list_intel_drivers

if __name__ == '__main__':
    print("Verificando Drivers!")
    # Get the drivers found and the quantity of drivers
    drivers_found, quantity_drivers = list_intel_drivers()
    # Parse the JSON result
    json_result = json.loads(drivers_found)


    # Imprime o resultado
    print(drivers_found)
    print("The number of drivers is:", quantity_drivers)