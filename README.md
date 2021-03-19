# TESTEN
Program for AE plus to export multiple schedules into .xlsx files. 

![Schermafbeelding 2021-03-10 142054](https://user-images.githubusercontent.com/39960806/111767832-780d8500-88a7-11eb-9355-dec7a0eebd0f.png)


This project started with a blank slate, there are a lot of possibilities as to how it is possible to export a bill of quantities. Understanding how to combine both programs with each other is not as straight forward as it seems. A project has products which are measured as numbers, meters or even full dimensions with length, width and height. Importing seems like the limitation, exporting out of Revit has more possibilities and working out the best method is the first problem. 


It quickly became clear that writing an own plugin seems like the best method. While this is not easy as other options such as using Dynamo, writing the plugin gives the most flexibility in the future and is something on which the company can keep building in the future. Reviewing the export methods is a big piece in the planning as reviewing and at the end choosing the optimal method is something on what the project keeps building and has to be right from the start.


Geo-IT is the limitation, it has a selected number of parameters which can be loaded in the bill of quantities. The exporting of Revit will be split up for each category of products. While the articles in numbers are exporting great and do not require a lot of tweaking.


While figuring out what is imported is important, figuring out what is not imported is as important as missing something in the calculation is even more dreadful. It became clear that getting an overview of what is not imported has to be implemented. Making the export even more valuable as mistakes can always happen. Testing is the biggest part of this project as it needs to be known how reliable it is. Knowing how the designer has to make the families and types correspondent to the export is crucial in succeeding. 

![Schermafbeelding 2021-03-19 113127](https://user-images.githubusercontent.com/39960806/111767160-b22a5700-88a6-11eb-9a69-8d9872a694cd.png)

These are the files which need to exported to an excel. 

For this setup we choose to split up each export from the start. There are 3 large bill of quantities for each project. Electricity, HVAC and plumbing. For importing each of these we need to split it up again, Geo-IT does not handle exceptions well so everything is arranged in schedules which count, dimensions and length. Creating these in panels and adding the right buttons to them will be added in the ExternalApplication. 

Each export has "2 modes" an overview with every Geo-IT link filled in and every schedule is a worksheet. The other export method gives everything filled in 1 worksheet and that workbook has another worksheet with everything blank to know what is not added.
