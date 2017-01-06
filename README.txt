============

pypowerworld

============

pypowerworld provides a user-friendly interface to the Powerworld COM object.

INPUT: The parameters for the methods and the object attributes are labeled identically to the Powerworld in the Auxiliary File documentation. 

OUPUT: Data output is sent in data frame format and set to the ‘output’ attribute.

ERROR HANDLING: If an error occurs, the ‘error’ attribute is set to True and the ‘error_message’ attribute is set to an error string value explaining the error. If an error is not encountered, the ‘error’ attribute is set to False.

TYPICAL IMPLEMENTATION:

	import pypowerworld

	case_name = ‘somecasename’
	pw = pypowerworld.open(case_name)
	pw.solve()

	if pw.error:
		print(“Error encountered solving power flow: %s” %pw.error_message)
		print(“Attempting to solve with immediate VAR limitations…”)
		pw.set_data()
		pw.solve()
		if pw.error:
			print(“The case could not be solved.”)
		else:
			print(“The case could be solved by immediate VAR limitations”)
	else:
		print(“The case could be solved”)