JSDoc2CHM
=========

This project converts documentation from JSDoc to a HTML Help Collection which can be integrated into Microsoft Document Explorer.

Essentially the project is just two VBScript files which create the HxT table of contents file and the HxF Include file that you need for a HTML Help 2 project.

![H2 Project Editor](http://www.williamsportwebdeveloper.com/images/H2-Project-Editor.jpg)

I was inspired to do this project by the [Famo.us](https://famo.us/) reference documentation which isn't provided in a Windows help file format. Although I could have gone to the trouble of adding it manually to my notes, I noticed that it would not be too difficult to do a conversion.

Therefore I have included the **Famous.HxS** file which can be integrated in Microsoft Document Explorer and the **Famous.chm** file which is a stand alone Windows help file.

Remember, the VBScripts can be used with any JSDoc output to produce the two files you'll need to convert your documentation into a standard Windows help file. 
