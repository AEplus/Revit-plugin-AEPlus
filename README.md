# AE Plus Revit API plugin
Program for AE plus to export multiple schedules into .xlsx files. 

![Schermafbeelding 2021-03-10 142054](https://user-images.githubusercontent.com/39960806/111767832-780d8500-88a7-11eb-9355-dec7a0eebd0f.png)


This project started with a blank slate, there are a lot of possibilities as to how it is possible to export a bill of quantities. Understanding how to combine both programs with each other is not as straight forward as it seems. A project has elements which are measured as numbers, meters or even full dimensions with length, width and height. Importing into Geo-IT seems like the limitation, exporting out of Revit has more possibilities and working out the best method is the first problem. With these two programs there is also “a third program”, these are the families which are provided from the manufacturers. These families or objects are also changing variables in connecting everything together for a bill of quantities.

It quickly became clear that writing an own plugin seems like the best method. While this is not easy as other options such as using Dynamo, writing the plugin gives the most flexibility in the future and is something on which the company can keep building in the future. Reviewing the export methods is a big piece in the planning as reviewing and at the end choosing the optimal method is something on what the project keeps building and must be right from the start. Starting in the right direction is a big part of succeeding as time is of essence. Choosing a plugin is riskier but allows more flexibility in the long run.

Geo-IT is the limitation, it has a selected number of parameters which can be loaded in the bill of quantities. The exporting of Revit will be split up for each category of products. Creating the schedules and Revit plugin to work around the limitations of Geo-IT are the way to go. Contacted Geo-IT for their view on what the goal is and how to achieve this and they suggested that my solution to achieve the automatic counting is fine.  
 
Figure 1: Import settings for Geo-IT software
As the main project will be a lot of code based, I opted to maintain it in a GitHub Repo (https://github.com/sbellinkx/Revit-plugin-AEPlus has been set to private). Following the Autodesk ‘your first plugin’ intro gave me the basic understanding of how to tackle this problem. This includes multiple branches to not get lost in little changes which can break the whole program. 
The first item on the list was the Revit Plugin, following the Autodesk “My First Plugin” https://knowledge.autodesk.com/support/revit-products/learn-explore/caas/simplecontent/content/my-first-revit-plug-overview.html
Adding the OpenOfficeXML library is needed to get the Excel export working. The number of Geo-IT must be the first field/row in the schedule. Excel sheets cannot extend 31 characters, this has been made dummy proof in the code however it has to be different for each schedule and cannot contain the same name. 
 
•	AE M52 Riolering materiaal zaaglijst isolatie 
•	AE M52 Riolering materiaal zaaglijst hulpstukken 

Are for exporting the same name and need to be different in the first 31 characters of the name. 
This was the result after two weeks of working on the plugin, researching the best options on how to code the plugin. Doing everything in Revit is almost impossible because the counting and lengths get mixed up. Splitting it up for each BOQ and separating lengths and counting objects is necessary. And multiple schedules will get duplicates to overcome this problem.
While looking around in the project everything seemed logically except cable trays was something which stood out because of the combination with width and height. Not knowing the full power and limitations of Revit I headed to the forum and askes my question there. https://forums.autodesk.com/t5/revit-mep-forum/add-value-to-each-modifyable-pipe/td-p/10142249

My question got answered and a calculated value was needed to do what I wanted to achieve. 
https://knowledge.autodesk.com/support/revit-products/learn-explore/caas/CloudHelp/cloudhelp/2017/ENU/Revit-Model/files/GUID-44EC6734-D36B-4083-8FCD-833B166D6C0A-htm.html
https://knowledge.autodesk.com/support/revit-products/learn-explore/caas/CloudHelp/cloudhelp/2017/ENU/Revit-Model/files/GUID-A0FA7A2C-9C1D-40F3-A808-73CD0A4A3F20-htm.html
While figuring out what is imported is important, figuring out what is not imported is as important as missing something in the calculation is even more dreadful. It became clear that getting an overview of what is not imported must be implemented. Making the export even more valuable as mistakes can always happen. Testing is the biggest part of this project as it needs to be known how reliable it is. Knowing how the designer must make the families and types of correspondent to the export is crucial in succeeding. Choosing to load every .csv file in for each line and then getting the index [0] so that it compares every 

The change was made to make 2 excels out of each schedule. One for an overview and one to import to Geo-IT. Each schedule will look the same, they all have the Geo-IT number in the first column which will be either empty or filled in. The program checks if the cell is filled in or not and copy this whole row to the corresponding excel file. 
 
After finishing and testing how this works, the next step is to finish the best way to get every Geo-IT number into the schedule. The way this works now is trying to implement an extra shared parameter which will be used to give each different type a number and combining it with the parameters in the schedule to figure out which Geo-IT number it is. While looking at every schedule a problem occurred. Revit does not allow to make calculated values through text, everything must be a number however the output can be text. 
While testing around with the numbers and text problems in Revit, everything will be in numbers because: 
Geo-IT Number: 60.52.40.11 is defined as text in Revit.
Revit number: 60524011 is defined as integer.

While this does include extra work, every element is now the same and this will get the Geo-IT Number through a calculated value in each schedule. 
Following this, the ductsystemtypes are the biggest problem for this. Not everything in the air schedules is the same and a type can be different to which air type system it is connected. The API provides a work around for this problem through the Dynamic Model Updater (DMU). The change scope for the updater is one of two things:
•An explicit list of element Ids in a document - only changes happening to those elements will trigger the Updater.
•An implicit list of elements communicated via an Element Filter - every changed element will be run against the filter, and if any pass, the Updater is triggered There are several options available for change type.

While figuring out how the updater works, it did not quite respond to every change that was made and had some problems with filling everything in. While struggling to find the solution two weeks later I discovered that the schedules will be filtered for each system type so making the updater is not necessary as you can add a separate number to each schedule which can be used for the calculated values. Each schedule will have different calculated values so this does not give problems. 
The next step was typing the calculated values file, I stumbled upon the problem that there is also a limit on how long a calculated value can be. This has now been split up for each schedule, this file is also maintained in a GitHub repository. https://github.com/sbellinkx/Calculatedvalues this repository is made public for the future if something is changed everyone from the company could just copy and paste it to their corresponding project in the meantime. 

Doing everything at once is too much to figure out at once. Anje made the decision to focus on the sanitary bill of quantities first because, these are only 1500 elements in comparison to HVAC which are 4500 elements. The roadblocks which are encountered here will be of smaller size and easily noticed when it is in a small project and containing a smaller number of elements.

 
In some schedules the calculated values stayed blank, the issue which is encountered here is that the sorting and grouping has some problems. When we itemize every instance of elements, we can see that the parameters of some of these elements are different in the same group. This is the issue as to why some stay blank, these are small modeler errors which are now figured out. In the future something like this will be detected quickly when returning in the blank excel worksheet. 
Now the final testing has begun, we can import the schedules straight into Geo-IT. Elements which are not imported get displayed as an error in the program. In Geo-IT we can export the whole project to an excel file and see how this compares to the manual counting.

Having contacted Geo-IT in the end for their new updates, there is an error with choosing the excel rows which straight out fails now, they added another option to pick the excel rows by name they have. This option does work but they made a remark that I can change the naming of the schedules so it automatically picks the right columns corresponding to the Geo-IT order. This a great option that makes it easier to import the counting. Combining length and counting is still not possible as these columns overlap each other. 

Compared to the manual counting of this project everything that could be filled in is correct and empty cells are displayed correct. In the end there will still be some problems in my program and how everything is set-up, but these will decrease if as the program will keep evolving and growing. Human errors can also be made and while using this workflow we can decrease the errors.

When looking back at this project, my weaknesses is a bit miss formulated, as the limited Revit knowledge is not only Revit but also how the company has setup their BOQ and how this is displayed in Revit. Problem solving skills are one of the most important skills, problems are inevitable and solving each problem to stumble on to the next one is what makes this interesting. Seeing a project grow and getting closer to the solutions is the drive to keep going. 

When looking back at my planning and time management I have invested more time than what initially was intended to. I sort of expected this because of the difficulty of the project and figuring everything out. I also spend 2 weeks on the dynamic updater which in the end was not necessary at all. I want to deliver something on which I am proud of and investing more time to achieve this is not a big obstacle for me.
