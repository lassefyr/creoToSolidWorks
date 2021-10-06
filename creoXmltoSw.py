# https://github.com/latefyr/kicadcreo.git
# MIT license
#
# Parse Creo xml to create Solidworks compatible Excel files
# 
#
# Copyright (C) LasseFyr 2021.
#
# This has been tested only with Creo 4.0 070
#
"""
    @package
    Generate a net list file.

    Command line: 
    python creoXmltoSw.py door_cables_creo.xml
		1. First parameter is the creo schematic .xml file created by kicad.
		
	Output:
		1. fileName+"from_to.xlsx"   		// From To Excel File
		2. fileName +"_sw_cbl.xml"			// Wires and Cables
		2. fileName +"_sw_comp.xlsx"		// Components
	
	
	
"""

from __future__ import print_function
from xml.dom import minidom
import sys
#import xlwt
import xlsxwriter
import os
import color_constants


class creoXmltoSw:
	def __init__(self):
		self.__infoString = ""
		self.__errorString = ""
		self.__warningString = ""	
		
	def getXmlItemValue( self, listToSearch, paramName, nameToFind ):			
		
		if( (isinstance(listToSearch, minidom.Element)) and (paramName)):		
			newList =  listToSearch.getElementsByTagName(paramName)
			for item in newList:
				if( item.getAttribute("name") == nameToFind ):
					return item.getAttribute("value")
							
		return ""
				

	def readConnections( self, fileName ):
		self.creoSchXmlName = fileName +".xml"
			
		print( "Convert Creo .xml to SW Excel\n" )								
		print( "-----------------------------\n" )								

		if( os.path.isfile(self.creoSchXmlName) ):
			print( "Kicad Schematic xml file found: = " + self.creoSchXmlName + "\n" )								
		else:
			print( "Creo Schematic xml NOT FOUND: " + self.creoSchXmlName + "\n" )
			print( "NOTE:You need to export Creo Schematic xml file with name: " + self.creoSchXmlName + "\n" )			
#				self.writeErrorStr( "Creo Schematic xml NOT FOUND: " + self.creoSchXmlName + "\n" )
#				self.writeInfoStr( "NOTE:You need to export Creo Schematic xml file with name: " + self.creoSchXmlName + "\n" )			
			return False
							
		creoXml = minidom.parse(self.creoSchXmlName)

		print( "Write Connections" )								
		print( "-----------------" )								
		
		connections = creoXml.getElementsByTagName("CONNECTION")
		self.WireName = "" 			# A descriptive identifier for the wire, such as T2 Signal or GND1. 
									# This field is for information only; the value is not used by the software.
		self.FromRef = ""			# The reference name for the first component to which the wire is connected, such as motor1 or con3. 
		self.FromPin = ""			# The pin number/name to which the wire is connected.
		self.FromPartNumber = ""	# The part number of the connector for this component reference, such as db9-plug or 5pindin-plug
		
		self.ToRef = ""				# 
		self.ToPin = ""				# 
		self.ToPartNumber = ""		# 
		
		self.CableName = ""			# If this wire is a core in a cable, the name of the cable. Leave blank for individual wires.
		self.CoreName = ""			# If this wire is a core in a cable, the name of the core in the cable. Leave blank for individual wires. 
		self.WireSpec = ""			# The part number of the wire or cable. 
		self.Other = ""				# Header names for user-defined attributes.			
		

		workbook = xlsxwriter.Workbook(fileName+"_from_to.xlsx")
		worksheet = workbook.add_worksheet("From To")
		worksheet.write(0, 0, 'Wire')
		worksheet.write(0, 1, 'Cable')
		worksheet.write(0, 2, 'Core')
		worksheet.write(0, 3, 'Spec')
		worksheet.write(0, 4, 'From Ref')
		worksheet.write(0, 5, 'Pin')
		worksheet.write(0, 6, 'Partno')
		worksheet.write(0, 7, 'To Ref')
		worksheet.write(0, 8, 'Pin')
		worksheet.write(0, 9, 'Partno')			
		worksheet.write(0, 10, 'Color')			

		print( 	"Wire,Cable,Core,Spec,From Ref,Pin,Partno,To Ref, Pin,Partno ")
						
		rowNumber = 1
		
		for connection in connections:
			tType 		= connection.getAttribute("type")
			tContext 	= connection.getAttribute("context")

			# Check whether cable header
			if( tType == "ASSEMBLY" and tContext == "NONE" ):
				self.CableName = connection.getAttribute("spoolID")[4:]
				self.WireSpec = self.CableName
				continue

			if( tType == "SINGLE" and tContext == "NONE" ):
				self.CableName = ""
				self.CoreName = ""

			if( tType == "SINGLE" and tContext == "CONNECTION" ): # THIS IS a CABLE
				self.CoreName =  "W"+(connection.getAttribute("name").split("_",1)[1])
			
			attachPar = connection.getElementsByTagName('ATTACH')
			self.WireName = connection.getAttribute("name")
			for param in attachPar:
				tFrom = param.getAttribute('node1ID')
				tTo = param.getAttribute('node2ID')
				
				#print("from = "+ tFrom + ", To = "+ tTo) 
				
				tempSplit = tFrom.split("_")
				self.FromRef = tempSplit[1]
				self.FromPin = "PIN_"+tempSplit[2]
							
				tComp = creoXml.getElementsByTagName("COMPONENT")
				for tComponent in tComp:
					if( tComponent.getAttribute("name") == tempSplit[1] ):
						self.FromPartNumber = tComponent.getAttribute("modelName")							
						
						portParams = tComponent.getElementsByTagName("PORT")

						compareParam = tempSplit[0]+"_"+tempSplit[1]+"_"+tempSplit[2]
						for portParam in portParams:																							
							tempList =  portParam.getElementsByTagName("SYS_PARAMETER")
							for item in tempList:
								if( item.getAttribute("id") == compareParam ):
									self.FromPin = self.getXmlItemValue( portParam, "PARAMETER", "ENTRY_PORT" )
						
		
				tempSplit = tTo.split("_")
				self.ToRef = tempSplit[1]
				self.toPin = "PIN_"+tempSplit[2]
				
				for tComponent in tComp:
					if( tComponent.getAttribute("name") == tempSplit[1] ):
						self.ToPartNumber = tComponent.getAttribute("modelName")			

						portParams = tComponent.getElementsByTagName("PORT")

						compareParam = tempSplit[0]+"_"+tempSplit[1]+"_"+tempSplit[2]
						for portParam in portParams:																							
							tempList =  portParam.getElementsByTagName("SYS_PARAMETER")
							for item in tempList:
								if( item.getAttribute("id") == compareParam ):
									self.toPin = self.getXmlItemValue( portParam, "PARAMETER", "ENTRY_PORT" )

								
				tSpoolName =  connection.getAttribute("spoolID")					
				tCabl = creoXml.getElementsByTagName("SPOOL")
				for tCable in tCabl:	
					if( tSpoolName.find(tCable.getAttribute("name"))!= -1 ):
						self.WireSpec = tCable.getAttribute("name")
		
				print( 	self.WireName + ", " +\
						self.CableName  + ", " +\
						self.CoreName  + ", " +\
						self.WireSpec  + ", " +\
						self.FromRef + ", " +\
						self.FromPin + ", " +\
						self.FromPartNumber + ", " +\
						self.ToRef  + ", " +\
						self.toPin  + ", " +\
						self.ToPartNumber  + ", " +\
						self.Other )

				worksheet.write(rowNumber, 0,  self.WireName)			
				worksheet.write(rowNumber, 1,  self.CableName)			
				worksheet.write(rowNumber, 2,  self.CoreName)			
				worksheet.write(rowNumber, 3,  self.WireSpec)			
				worksheet.write(rowNumber, 4,  self.FromRef)			
				worksheet.write(rowNumber, 5,  self.FromPin)			
				worksheet.write(rowNumber, 6,  self.FromPartNumber)			
				worksheet.write(rowNumber, 7,  self.ToRef)			
				worksheet.write(rowNumber, 8,  self.toPin)			
				worksheet.write(rowNumber, 9,  self.ToPartNumber)			
				worksheet.write(rowNumber, 10, self.Other)			
				rowNumber = rowNumber + 1
		'''
		while True:
			try:
				#workbook.close()
				workbook.save('from_to.xls')
			except xlwt.exceptions.FileCreateError as e:
				decision = input("Exception caught in workbook.close(): %s\n"
				"Please close the file if it is open in Excel.\n"
                "Try to write file again? [Y/n]: " % e)
				
				if decision != 'n':
					continue
			break	
		'''
		
		while True:
			try:
				workbook.close()
			except xlsxwriter.exceptions.FileCreateError as e:
				decision = input("Exception caught in workbook.close(): %s\n"
				"Please close the file if it is open in Excel.\n"
                "Try to write file again? [Y/n]: " % e)
				if decision != 'n':
					continue
			break
						
		# workbook.close()			
							
			# Put lenghts to Kicad Schematic
			# Do not overwrite the original file
			# self.writeInfoStr( "\nProcessing wires and cables:\n" )								
			# self.writeInfoStr( "----------------------------\n" )								
			# self.writeInfoStr( str(self.refDesVals) + "\n\n" )								
		

		# Spools Write Spool XML --------------------------------------------------------
		cableID = 1
		coreID = 1
		wireID = 1
		
		swColorValue = color_constants.RGB
			
		print( 	"Spools")
		print( 	"------")			
		
		rowNumber = 1
		
		self.partNo = ""
		self.cableName = ""
		self.NumberOfCores = ""
		self.wireName = ""
		self.description = ""
		self.diameter = ""
		self.awgNum = ""
		self.wireColor = ""
		self.displayColor = ""
		self.minBendRadius = ""
		
		
		actualFileOpened = False
		outputFileName = fileName +"_sw_cbl.xml"
		if len(outputFileName) > 2:
			try:
				if sys.version_info.major < 3:		
					fout = open(outputFileName, "w")
				else:
					fout = open(outputFileName, "w", encoding='utf-8')
				actualFileOpened = True
			except IOError:
				e = "Can't open output file for writing: " + outputFileName
				print( __file__, ":", e, sys.stderr )
				fout = sys.stdout
		else:
			fout = sys.stdout		
		
		print("<?xml version=\"1.0\" encoding=\"UTF-8\"?>", file = fout)
		print("<CableLibrary xmlns=\"http://www.solidworks.com/cablelibrary\">", file = fout)
		
		tCabl = creoXml.getElementsByTagName("SPOOL")
		for tCable in tCabl:
			self.partNo = tCable.getAttribute("name")
			self.cableName = self.partNo # tCable.getAttribute("name"))
			
			# This is a Cable header
			if( tCable.getAttribute("subType") == "CABLE_SPOOL" ):
				self.partNo = tCable.getAttribute("name")
				self.cableName = tCable.getAttribute("name")
				self.wireName = ""	
												  
				wireDescription = self.getXmlItemValue( tCable, "PARAMETER", "VENDOR_PN" )
				self.description = wireDescription
				self.diameter = self.getXmlItemValue( tCable, "PARAMETER", "THICKNESS" )
				self.awgNum = ""
				self.wireColor = self.getXmlItemValue( tCable, "PARAMETER", "COLOR" ).capitalize()
				self.displayColor = str(swColorValue.getIfromRGB(self.wireColor.lower()))
				self.NumberOfCores = self.getXmlItemValue( tCable, "PARAMETER", "NUM_COND" )
				self.minBendRadius = self.getXmlItemValue( tCable, "PARAMETER", "MIN_BEND_RADIUS" )
				
				print( 	"<!--Cable With Cores-------------------------------->")
				print( 	"<cable ID=\""+str(cableID)+"\" >", file = fout )				
				print( 	"<cableName value =\""+self.cableName+"\" />", file = fout )
				print( 	"<partNumber value =\""+self.partNo+"\" />", file = fout  )
				print( 	"<Description value =\""+self.description+"\" />", file = fout )
				print( 	"<outerDia value =\""+self.diameter+"mm\" />", file = fout )
				print( 	"<color value =\""+self.wireColor+"\" />", file = fout )
				print( 	"<SWColor value =\""+self.displayColor+"\" />", file = fout )
				print( 	"<NoOfCores value =\""+self.NumberOfCores+"\" />", file = fout )
				print( 	"<MinBendRadius value =\""+self.minBendRadius+"mm\" />", file = fout )
				print( 	"<MassperUnitLength value =\"\" />", file = fout )				
				
				wireNum = 1
				coreID = 1
				for tInline in tCabl: # find sub wires
					if( tInline.getAttribute("type") == "INLINE_SPOOL" ):
						subWireName = tInline.getAttribute("name")
																				
						if( subWireName.find(self.cableName) != -1 ):
							self.partNo = ""
							#self.cableName = tCable.getAttribute("name"))
							self.wireName = "W"+ str(wireNum)	
							wireNum = wireNum+1
							# wireDescription = self.getXmlItemValue( tCable, "PARAMETER", "VENDOR_PN" ):
							# self.description = wireDescription
							self.diameter = self.getXmlItemValue( tInline, "PARAMETER", "THICKNESS" )
							self.awgNum = ""
							self.wireColor = self.getXmlItemValue( tInline, "PARAMETER", "COLOR" ).capitalize()
							self.displayColor = str(swColorValue.getIfromRGB(self.wireColor.lower()))
							
							
							print( 	"<core ID=\""+str(coreID)+"\" >", file = fout )
							coreID = coreID+1
							print( 	"   <AdditionalProperty ID=\"1\">", file = fout )
							print( 	"    <PropertyName value=\"Cablename\"/>", file = fout )
							print( 	"    <PropertyValue value=\""+self.cableName+"\"/>", file = fout )
							print( 	"   </AdditionalProperty>", file = fout )							
							print( 	"   <coreName value =\""+self.wireName+"\" />", file = fout )
							print( 	"   <conductorSize value =\"Unset\" />", file = fout )
							print( 	"   <conductorOD value =\""+self.diameter+"mm\" />", file = fout )
							print( 	"   <color value =\""+self.wireColor+"\" />", file = fout )
							print( 	"   <SWColor value =\""+self.displayColor+"\" />", file = fout )
							print( 	"   <MinBendRadius value =\"1mm\" />", file = fout )
							print( 	"   <Seal value =\"Unset\" />", file = fout )
							print( 	"   <MassperUnitLength value =\"Unset\" />", file = fout )							
							print( 	"</core>", file = fout )							
					
				print( 	"</cable>", file = fout )
				cableID = cableID + 1
				

									
			# This is a single conductor				
			if( tCable.getAttribute("type") == "NORMAL_SPOOL" and tCable.getAttribute("subType") == "WIRE_SPOOL" ):
				self.partNo = tCable.getAttribute("name")
				self.cableName = ""
				
			# All connections that use this wire...
				# wireNamesUsingWire = creoXml.getElementsByTagName("CONNECTION")
					
				#for wire in wireNamesUsingWire: # find sub wires
				#	if( wire.getAttribute("spoolID")[2:] == self.partNo ):
				#		self.wireName =  wire.getAttribute("name")
				
				self.wireName = self.partNo				
				wireDescription = self.getXmlItemValue( tCable, "PARAMETER", "VENDOR_PN" )
				self.description = wireDescription
				self.diameter = self.getXmlItemValue( tCable, "PARAMETER", "THICKNESS" )
				self.awgNum = ""
				self.wireColor = self.getXmlItemValue( tCable, "PARAMETER", "COLOR" ).capitalize()
				self.displayColor = str(swColorValue.getIfromRGB(self.wireColor.lower()))
				self.minBendRadius = self.getXmlItemValue( tCable, "PARAMETER", "MIN_BEND_RADIUS" )
				
				print( "<!--Single Wire ------------------------------------>")
				print( "<wire ID=\""+str(wireID)+"\" >", file = fout )			
				wireID = wireID + 1
				print( "   <WireName value =\""+self.wireName+"\" />", file = fout )
				print( "   <partNumber value =\""+self.partNo+"\" />", file = fout )
				print( "   <Description value =\""+self.description+"\" />", file = fout )
				print( "   <WireOD value =\""+self.diameter+"mm\" />", file = fout )
				print( "   <color value =\""+self.wireColor+"\" />", file = fout )
				print( "   <SWColor value =\""+self.displayColor+"\" />", file = fout )
				print( "   <MinBendRadius value =\""+self.minBendRadius+"mm\" />", file = fout )
				print( "   <Size value=\"\" />", file = fout )
				print( "   <Seal value=\"\" />", file = fout )
				print( "   <MassperUnitLength value =\"\" />", file = fout )
				print( "</wire>", file = fout )

		print(  "</CableLibrary>", file = fout )
		if actualFileOpened == True:
			fout.close( )
		
		
		print( 	"Components")
		print( 	"----------")
	

		components = creoXml.getElementsByTagName("COMPONENT")
		self.partNo = "" 				# Part name in from-to list
		self.libName = ""				# Solidworks Component name
		self.configName = "Default"		# The pin number/name to which the wire is connected.
		self.Description = ""			# Description of the part (connector, ring terminal, other reference, such as db9-plug or 5pindin-plug
					
		self.componentList = []
		
		print( 	"Partno, Libname, Conmfigname, Descrioption ")
		
		workbook = xlsxwriter.Workbook(fileName +"_sw_comp.xlsx")
		worksheet = workbook.add_worksheet("Connectors")		
		worksheet.write(0, 0, 'Partno')
		worksheet.write(0, 1, 'Libname')
		worksheet.write(0, 2, 'Configname')
		worksheet.write(0, 3, 'Descrioption')
		rowNumber = 1
						
		for component in components:
			self.partNo = component.getAttribute("modelName")								
			self.libName = component.getAttribute("modelName")+".sldprt"								
			self.Description = ""
			
			if ( self.partNo in self.componentList ):
				continue

			self.componentList.append(self.partNo)				
			self.Description = self.getXmlItemValue( component, "PARAMETER", "OBJ_TYPE" )
									
			print( 	self.partNo + ", " +\
					self.libName  + ", " +\
					self.configName  + ", " +\
					self.Description )

			worksheet.write(rowNumber, 0, self.partNo)
			worksheet.write(rowNumber, 1, self.libName)
			worksheet.write(rowNumber, 2, self.configName)
			worksheet.write(rowNumber, 3, self.Description)
			rowNumber = rowNumber + 1

		while True:
			try:
				workbook.close()
			except xlsxwriter.exceptions.FileCreateError as e:
				decision = input("Exception caught in workbook.close(): %s\n"
				"Please close the file if it is open in Excel.\n"
                "Try to write file again? [Y/n]: " % e)
				if decision != 'n':
					continue	
			break
		'''
		while True:
			try:
				workbook.save('components.xls')
			except IOError:
				decision = input("Exception caught in workbook.close(): %s\n"
				"Please close the file if it is open in Excel.\n"
                "Try to write file again? [Y/n]: ")
				
				if decision != 'n':
					continue
			break
		'''					
#-----------------------------------------------------------------------------------------
# If this is called Independently
#
# Create instance and call with parameters
#
#-----------------------------------------------------------------------------------------				
if __name__ == '__main__':      
	fileToProcess = sys.argv[1]    				# unpack 2 command line arguments  
	
	# Split the file extension away if it exists
	fileToProcess = os.path.splitext(fileToProcess)[0]	
	
	creoToSw = creoXmltoSw( )
	creoToSw.readConnections( fileToProcess )
	
	print("DONE")
#	print("Info", file=sys.stdout)
#	print( creoCablelengths.getInfoStr(), file=sys.stdout )

#	print("Warnigns", file=sys.stdout)
#	print( creoCablelengths.getWarningStr(), file=sys.stdout )

#	print("Errors", file=sys.stderr)
#	print( creoCablelengths.getErrorStr(), file=sys.stderr )
#	print( "Please Reload the Kicad Schematic", file=sys.stdout )