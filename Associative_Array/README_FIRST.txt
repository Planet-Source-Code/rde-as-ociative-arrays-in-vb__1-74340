
Remember that the Collection object is
already compiled so for a fair comparison:

Open the project group TEST.vbg
In project explorer select AssocAVL.vbp
Compile the Dll
 If you get a compatibility message:
   Compile anyway (again) then
   Enter project properties > Component tab
   Reset project to binary compatibility
   Compile again
In project explorer select Tests.vbp
Select Project menu > References
 Uncheck/Check the Associative entry
 Compile the tests project

Run Tests.exe
