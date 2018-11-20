# Automatic folder creation

Script purpose : Automatically create folders for each Ansys APDL study listed in the "tests_combination.xlsx".

The use of this kind of script could be useful for optimization problem for example.


## Setting up

1- "tests_combination.xlsx" --> "path" : add the Ansys installation path
2- "tests_combination.xlsx" --> "tests" : add the list of tests


## Launching Python script

Launch the "main.py"

```
In a command prompt : python main.py
```


## Results

1- Creation of each folder
2- Creation of a "variables.mac" file in each folder
3- Creation of a "lancement_ansys.bat"


```
Do not forget to import the "variables.mac" in your "main.mac" file !
```


## Launching Ansys

Launch the "lancement_ansys.bat".
Each study should should be launched automatically