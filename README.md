# TESTEN
Program for AE plus to export multiple schedules into .xlsx files. 

This project started with a blank slate, there are a lot of possibilities as to how it is possible to export a bill of quantities. Understanding how to combine both programs with each other is not as straight forward as it seems. A project has products which are measured as numbers, meters or even full dimensions with length, width and height. Importing seems like the limitation, exporting out of Revit has more possibilities and working out the best method is the first problem. 


It quickly became clear that writing an own plugin seems like the best method. While this is not easy as other options such as using Dynamo, writing the plugin gives the most flexibility in the future and is something on which the company can keep building in the future. Reviewing the export methods is a big piece in the planning as reviewing and at the end choosing the optimal method is something on what the project keeps building and has to be right from the start.


Geo-IT is the limitation, it has a selected number of parameters which can be loaded in the bill of quantities. The exporting of Revit will be split up for each category of products. While the articles in numbers are exporting great and do not require a lot of tweaking.


While figuring out what is imported is important, figuring out what is not imported is as important as missing something in the calculation is even more dreadful. It became clear that getting an overview of what is not imported has to be implemented. Making the export even more valuable as mistakes can always happen. Testing is the biggest part of this project as it needs to be known how reliable it is. Knowing how the designer has to make the families and types correspondent to the export is crucial in succeeding. 
