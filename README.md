# creoToSolidWorks

This project is related to the kicadToCreo. You can design you wiring diagram by using Kicad. From the Kicad Schematic you can 
create the creo schematic compatible xml netlist. 

You can then use this netlist to generate the below mentioned files for solidworks.
It seems that you only need the from-to and the sw_cbl.xml files to create wires in Solidworks.

Converts the Kicad Creo output to Solidworks FromTo file

This was created to test drive the SolidWorks premium. We used an existing Creo design and
created the necessary files for solidworks.

The Script creates following files
* fileName+"from_to.xlsx"     // From To Excel File
* fileName +"_sw_cbl.xml"     // Wires and Cables
* fileName +"_sw_comp.xlsx"   // Components

Solidworks routes wires somewhat differently from Creo.

After initial trial it seems that (todo)
- Perhaps cables and wires should be appended to existing cable-wire xml file?
- Additional parameters needed?
- 

I do not know how to create multiple cable asseblies from a single from-to excel file.

Connector mating (single coordinate) was something I did not succeed in without selecting each<br>
connector separately and telling it to "Aling Axes". 

I will continue this project when/if we start to adopt SW more in our desings.

As always... If your company has money then please use
* Creo Schematic for Creo
* Solidworks Electrical for Solidworks
