XML-VB-Test-Data-Generator
==========================

An XML driven Test Data Generator written in VB6

Compendium Developments Test Data Generator Prototype
Compendium-TDG
copyright Compendium Developments 2005

Alpha code - no warranty.

Overview
--------
Compendium-TDG is a test data generator using input XML files.

Minimal instructions
--------------------

The application requires VB runtimes in order to run. These are not supplied.
The application also requires MSXML v3 or above. Also not supplied, but if you use
a recent version of IE then MSXML should be installed.

Also supplied is a sample xml file which describes the format of the input files
and has some example rules.

To Use the GUI, run the app.

press [...] and select an input xml file to process.

Then press [Parse XML]...
The input xml file will be parsed and examples of the
rules in the file will be displayed in the text box.
Any found OutputRules will be displayed in the list box.

Double click an output rule to run the rule and send the output to a chosen file.
All records are crlf terminated


Basic Concepts
--------------
Sets are sets of data elements
Rules act on sets or other rules to create data.
OutputRules are used to generate records in an output file.

The example data
----------------
The example data.xml file, the output rule when run will create 100 records in
csv format, where:
 the first column is an item from the User Role set, 
 the second column is an item from the Customer Role set,

Run from the command line
-------------------------
Compendium-TDG can also be run from the command line e.g.
testDataGenerator.exe -i c:\data\data.xml -o User Roles -of c:\output\userroles.csv

-i defines the input xml file
-o the output rule to execute
-of the output file

-o and -of can be repeated so

testDataGenerator.exe -i c:\data\data.xml -o User Roles -of c:\output\userroles.csv -of c:\output\userroles2.csv
would rule the same output rule from the same input file but produce a differnt output file

if -o is repeated then the output file variable is blanked and has to be defined again by calling -of again
testDataGenerator.exe -i c:\data\data.xml -o User Roles -of c:\output\userroles.csv -o output rule2 -of c:\output\userroles2.csv

if the argument -s is used then no msgbox describing the arguments is given


XML Format documentation
------------------------
The xml formatting rules can be found in the data\data.xml file and have been copied in here for initial distributions.

```
<!-- DataDefinitions
	contains Sets and Rules and OutputRules

	Sets contains
		Set (with a mandatory name) contains
			Element with the details in the contents
			
		A can also be stored in a file
		 <set name = "aValue" fileName="rowdata.txt" fileType="ROWS"/>
		 ' filename can be absolute e.g. c:\files\data.txt, or is assumed in the same dir as the xml file
		 file is loaded in to memory once during processing when the set is accessed by a rule
		 a ROWS file has a single value on each row

		 <set name = "csv2Value" fileName="csvdata.csv" fileType="CSVCOLS" colID="2"/>
		 a CSVCOLS file is a CSV file, and the colID specifies the column ID of the data
		 default columnID is 1

	Rules contains
		Rule (with a mandatory name)

		a rule can contain any of the rule blocks in any order

		Rule Blocks:

			SetOperation
			Range
			Term
			Optional
			Choice
			Repeat

	OutputRules contains
		OutputRule (with a mandatory name) and optional type

		an outputRule can contain any of the rule blocks in any order
		but only <record></record> blocks are written out during the outputRule processing
		
		type can be TEXT, XML or CSV
		
		when 
		 <OutputRule name="asText" type = "TEXT">
		 	 an records are written as is, and fields are basically ignored, although all
		 	 contents are written out
		 </OutputRule>

		 <OutputRule name="asXML" type = "XML">
		 	records must have names and fields are used and must have names
		 	written out as XML format
		 	 <recordname>
		 	   <afieldname>fieldvalue</afieldname>
		 	 </recordname>
		 </OutputRule>

		 <OutputRule name="asCSV" type = "CSV">
		 	records are written out with fields comma delimited and fields quoted where necessary
		 </OutputRule>


		Rule Blocks Allowed:

			SetOperation
			Range
			Term
			Optional
			Choice
			Repeat

		OutputRule Blocks:
			Record


		SetOperation
		============

		A SetOperation is a rule for combining sets.

		A SetOperation has a type which can be {Union,Intersection,Difference}

		A SetOperation can contain any number of OperatesOn which document the sets which
		the SetOperation block operates on.

		The rule will return a single value from the set resulting from the operation
		on the OperatesOn sets.

		Range
		=====

		A Range block is a rule for specifing a range of information.

		A Range block has a type which determines what kind of range it is {int,date,char}

		An int range is a number 'from' some value, 'to' some other value. It can have a specified:
		. return 'width' which can be
		. 'paddedWith' some character and
		. 'padded' from some direction {Left,Right}

		A date range is a range of dates 'from' some date 'to' some other date. It can have a specified:
		. 'format' which determines what is returned - default is 'dddddd ttttt'

		A char range is a range of characters 'from' some char 'to' some other char.

		A range block returns a single value from the range.

		Term
		====
		A term is a simple way of getting information into the rule, it could be a literal, contained
		in the body of the term, or it could be a 'name'd rule or set.
		 <Term>X</Term> to get a literal X
		 <Term name="aname"/> to use the value of a rule 'aname' or a set 'aname' or a variable 'aname'
		   the names are checked for matches in the order of Rule, Set, Variable
		   variables are defined with the <AS name="aname"></AS> construct

		Optional
		========

		A block which is optional i.e. defaults to a 50-50 chance of appearing or not.
			attribute probability is a double between 0 and 1, 50-50 = 0.5 which is the default

		Can contain other blocks.

		Choice
		======

		Contains <Option> blocks which have an optional weighting attribute
		A choice of blocks. One of the contained Option blocks will be selected. Option blocks can
		be weighted to have a higher (or lesser) chance of being chosen.
			e.g. <option weighting="2"> is twice as likely to be chosen as an option
			     block with a weighting of 1
			     if no weighting is provided then it defaults to 1

		Repeat
		======
		A repeat block is a block where the contents of the block are repeated a defined or random
		number of times.

		'from' the minimum number of times to repeat
		'to' the maximum number of times to repeat
			if no to is provided then the default is used.  (default currently = 100)
		'fixed' = "true" or "false" - false to get a random number in the repeat range
		   defaults to fixed
		   e.g. if we have <Repeat from="1"> then this will create a repeat
			                  loop from 1 to the default of 100
			                 <repeat from="1" fixed="false> then this will create a repeat
			                  loop of from 1 to a random value between 1 and 100
			                 <Repeat from="1" to "50">  a repeat loop from 1 to 50
			                 <Repeat from="1" to "50" fixed="false">
											  a repeat loop from 1 to a random value between 1 and 50

		Record
		=======
		A record block outputs the whole contents in the form of a record,
		e.g. with a newline at the end in a text file
		
		record has an optional name, the name is used when the record is output as part
		of an OutputRule with type XML
			<Record name="users">

			optional attribute csvheader which can be set to "TRUE" or "FALSE"
			when the record is used in an outputRule and format is CSV, this record will be
			output with a header or all the FieldNames, this will be performed once
			per outputRule  e.g. <Record name="users" csvheader="FALSE">
			
				you might also want to consider using the CSVHEADER function which works in
				outputRules to separate data from format, instead of the csvheader attribute.
				see the CSVHEADER function for more details


		Field
		=====
		A number of Field blocks can be nested within a Record block to aid creation
		of output records and easily switch the format between CSV or XML
		

		Functions
		=========

			<GUID/> has no children, will return a GUID optional attribute braces="YES"
				so <GUID braces="YES"/> would return {xxxxxx-xxxx-etc.}
				braces defaults to NO and just returns the GUID xxxxxx-xxxx-etc.


			<TRIM>
				remove any white space before or after the enclosed value
			</TRIM>
			  optional attribute border="LR" or "L" or "R"
			  default functionality is LR to trim from both sides
			  but by doing just L you just trim the left and for R just the right


			<PADSTRING maxlen="20" border="L" with="*">
				anything within the PADSTRING tags will be padded out to maxlen chars
				 using the 'with' character
				if it is already longer than maxlen then it will be truncated to maxlen
				can be padded on the left (border="L") or the right (border="R")
			</PADSTRING>
				mandatory 'maxlen' for the string
				optional 'border' can be L or R - defaults to L
				optional 'with' as the pad character - defaults to space


			<RIGHT len="20">
				return the number of characters stated by len from the right of the string
			</RIGHT>

			<LEFT len="20">
				return the number of characters stated by len from the left of the string
			</LEFT>

			<SUBSTR start="2" len="20">
				return the number of characters stated by len from character number start
			</LEFT>

			<AS name="aname">
				create the results of the enclosed items as a variable called name
				if a name is reused then the value is overwritten
				if the name is the same as a rule or a set, then the rule or set will be
				brought back when used in a <TERM> block
				variables are global once defined
			</AS>
			
			<CSVHEADER show="TRUE">
				show can be "TRUE" or "FALSE" - default is "FALSE"
				when this is wrapped around a record which is being output in a CSV outputRule
				then it can control if the header (built from fieldNames) is shown (once per OutputRule)
				or not
			</CSVHEADER>

```


Licensed Under Apache 2.0
--------------------------


Copyright 2005 Alan Richardson - Compendium Developments

   Licensed under the Apache License, Version 2.0 (the "License");
   you may not use this file except in compliance with the License.
   You may obtain a copy of the License at

     http://www.apache.org/licenses/LICENSE-2.0

   Unless required by applicable law or agreed to in writing, software
   distributed under the License is distributed on an "AS IS" BASIS,
   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
   See the License for the specific language governing permissions and
   limitations under the License.