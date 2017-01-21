======================

==== pypowerworld ====

======================


==== OVERVIEW ====

pypowerworld provides a user-friendly interface to the Powerworld COM object. It instantness as a PowerWorld class object with a case. If the case fails to open, the program does not error out, but prints an error and proceeds without a case. The ‘open_case’ method may be used to attempt to open the PowerWorld case again.


==== REQUIREMENTS ====

This requires COM which means it must be run on a Windows OS. It also requires the user to have PowerWorld and SimAuto licenses. Any PowerWorld tools which require additional license are also necessary.


==== INPUT ====

The inputs are labeled identically to the Powerworld Auxiliary File documentation. The inputs can either be method parameters or attributes used to set data or get data from the PowerWorld program. 


==== OUPUT ====

Data output is sent in data frame format and set to the ‘output’ attribute. This is true only when the get data method is run. For all other operations, this attribute will be set to None. 


==== ERROR HANDLING ====

If an error occurs, the ‘error’ attribute is set to True and the ‘error_message’ attribute is set to an error string value explaining the error. If an error is not encountered, the ‘error’ attribute is set to False and the ‘error_message’ is set to an empty string.


==== TYPICAL IMPLEMENTATION ====

	import pypowerworld

	case_path = r‘somepath/somecase.pwb’
	pw = pypowerworld(case_path)
	pw.solve()

	if pw.error:
		print(“Error encountered solving power flow: %s” % pw.error_message)
	else:
		print(“The case was solved using the Newton-Raphson method.”)

