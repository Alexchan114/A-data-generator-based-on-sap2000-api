import comtypes.client
import random
import numpy as np
import os

############################ Introduction ##############################
# Run this code while running SAP2000 project
# Sap2000 V23.3.1 is recommended
# You can find every api used in this program in CSI_OAPI_Documentation, 
# which is in the SAP2000 installation directory
# Version: 1.0
########################################################################




# Create directories to store the data and labels
CirFre_Data = 'CirFre_Data'
if not os.path.exists(CirFre_Data):
    os.makedirs(CirFre_Data)

CirFre_Label = 'CirFre_Label'
if not os.path.exists(CirFre_Label):
    os.makedirs(CirFre_Label)   



# Connect to the running instance of SAP2000
sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")
SapModel = sap_object.SapModel

# Unlock the sap2000 model
SapModel.SetModelIsLocked(False)

# Set the model units to N/mm^2
SapModel.SetPresentUnits(9)  # 9 for N/mm^2

#Randomly distribute goods in the warehouse
def random_goods_distribution(num_z, num_x, num_y):
    # create a num_z * num_x * num_y matrix to store the distribution of goods
    distribution_matrix = []
    
    for layer in range(num_z):
        # Randomly distribute goods in each layer,1 for goods, 0 for no goods
        distribution = [random.choice([0, 1]) for _ in range(num_x * num_y)]
        distribution_matrix.append(distribution)  # append the distribution of each layer
        
    return distribution_matrix

############# configurations #############
num_trials = 5 # epoch for data generation
surface_pressure_area = 0.001  # pressure for each area
num_z = 3 #  3 for real , 3 for test
num_x = 11 # 11 for real , 7 for test
num_y = 14 # 14 for real , 7 for test


#Set area load  for each area according to the goods distribution matrix and run the analysis and design
num_loc = 462   # Total goos position, storage&AGVway included (147 for test   462 for real)
for trial in range(num_trials):
    # Randomly distribute goods in the warehouse
    distribution_matrix = random_goods_distribution(num_z, num_x, num_y)
    
    #print(f"Trial {trial + 1}: Goods Distribution Matrix:")
    #for layer in distribution_matrix:
        #print(layer)

    flattened_distribution = [item for sublist in distribution_matrix for item in sublist]

    # Set the area load for each area according to the goods distribution matrix
    for loc in range(num_loc):
        if flattened_distribution[loc] == 1:
            area_name = str(loc + 1)
            
            ret=sap_object.SapModel.AreaObj.SetLoadUniform(area_name,
                                                       "DEAD",
                                                       surface_pressure_area,
                                                       10,
                                                       True,
                                                       "Global",   #"local"
                                                       0)
            #ret = sap_object.SapModel.AreaObj.SetLoadUniform(area_name,"DEAD",surface_pressure_area,10,True,"Global",0)
            #if ret != 0:
                #print("Area load not assigned successfully.")
            #else:
                #print("Area load assigned successfully.")
            

    # Run the analysis
    ret=SapModel.Analyze.RunAnalysis()

    #ret = SapModel.Analyze.RunAnalysis()
    #if ret != 0:
        #print("Analysis did not run successfully.")
    #else:
        #print("Analysis run successfully.")

    NumberResults = 12
    LoadCase = []
    StepType = []
    StepNum = []
    Period = []
    Frequency = []
    CircFreq = []
    EigenValue = []

######################### Get first 12 circ_frequences of the structure ####################
    
    result = SapModel.Results.ModalPeriod(NumberResults, 
                                          LoadCase, 
                                          StepType, 
                                          StepNum, 
                                          Period, 
                                          Frequency, 
                                          CircFreq, 
                                          EigenValue)
    #print(result[6])  # 6 for CircFreq

    np.save(os.path.join(CirFre_Data,f'data_cirfre_{trial}'), distribution_matrix)
    np.save(os.path.join(CirFre_Label,f'label_cirfre_{trial}'), result[6])

     #Unlock the model
    SapModel.SetModelIsLocked(False)


    # Delete the analysis results
    ret=sap_object.SapModel.Analyze.DeleteResults(" ", True)

    #ret = sap_object.SapModel.Analyze.DeleteResults(" ", True)
    #if ret != 0:
        #print("Results not deleted successfully.")
    #else:
        #print("Results deleted successfully.")

    # Delete the area load for each area
    for load in range(num_loc):
        if flattened_distribution[load] == 1:
            load_name = str(load + 1)

            ret=sap_object.SapModel.AreaObj.DeleteLoadUniform(load_name, 
                                                          "DEAD", 
                                                          0)

            #ret= sap_object.SapModel.AreaObj.DeleteLoadUniform(load_name, "DEAD", 0)
            #if ret != 0:
                #print("Area load not deleted successfully.")
            #else:
                #print("Area load deleted successfully.")

    #Unlock the model again, in case the model is locked
    SapModel.SetModelIsLocked(False)
    


print(" ALL process done ")

# Close SAP2000 
#sap_object.ApplicationExit(False)