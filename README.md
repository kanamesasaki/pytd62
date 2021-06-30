# pytd62

pytd62 provides a set of functions for Thermal Desktop users to handle a thermal model more easily via python. 
Thermal Desktop is a thermal analysis software developed by C&R Technologies. Thermal Desktop version 6.2 offers programmable interface OpenTD, which allows users to automate many of the model development and analysis tasks. OpenTD is best communicated by using a .NET language such as C#. However, it can be accessed from python as well. 
OpenTD allows users to access detailed parameter settings and control features. However in some cases, it may not be convenient to specify all the details by writing a long script. Therefore, it would be useful to have some wrapper functions for operations which are repetitively needed during the thermal model development and analysis. For example, pytd62 provides functions for the following functionalities: 

* Create surface elements, solid elements, and nodes from a CSV format input file.
* Define material properties from a CSV format input file.
* Define node-to-node conductors from a CSV format input file.
* Calculate area, volume, and heat capacity of the elements.

For more details of pytd62 functionalities, please visit the pytd62 web page:
https://kanamesasaki.github.io/pytd62/