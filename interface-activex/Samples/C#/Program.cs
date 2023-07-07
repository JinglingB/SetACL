using System;
using SETACLLib;

//
// Example program showing how to use the OCX version of SetACL from C#
//
// Author:       Helge Klein
// Tested with:  Visual Studio 2010, .NET 4
//
// Requirements: 1) SetACL.ocx must be present and registered (with regsvr32)
//               2) A reference to SetACL must be added as a COM reference to the Visual Studio project
//
namespace SetACLFromCSharp
{
	class Program
	{
		static int Main(string[] args)
		{
			// Create a new SetACL instance
			SetACL setacl = new SetACL();

			// Attach a handler routine to SetACL's message event so we can receive anything SetACL wants to tell us
			setacl.MessageEvent += new _DSetACLEvents_MessageEventEventHandler(MessageHandler);

			// Set the object to process
			int retCode = setacl.SetObject(@"C:\Test", (int) SE_OBJECT_TYPE.SE_FILE_OBJECT);
			if (retCode != (int) RETCODES.RTN_OK)
				return 1;

			// Set the action (what should SetACL do?): add an access control entry (ACE)
			setacl.SetAction((int) ACTIONS.ACTN_ADDACE);
			if (retCode != (int)RETCODES.RTN_OK)
				return 1;

			// Set parameters for the action: Everyone full access
			setacl.AddACE("Everyone", false, "Full", 0, false, (int) ACCESS_MODE.SET_ACCESS, (int) SDINFO.ACL_DACL);
			if (retCode != (int)RETCODES.RTN_OK)
				return 1;

			// Add another action: configure inheritance from the parent object
			setacl.AddAction((int)ACTIONS.ACTN_SETINHFROMPAR);
			if (retCode != (int)RETCODES.RTN_OK)
				return 1;

			// Set parameters for the action: block inheritance for the DACL, leave the SACL unchanged
			setacl.SetObjectFlags((int) INHERITANCE.INHPARNOCOPY, (int) INHERITANCE.INHPARNOCHANGE, true, false);
			if (retCode != (int)RETCODES.RTN_OK)
				return 1;

			// Now apply the settings (do the actual work and change permissions)
			setacl.Run();
			if (retCode != (int)RETCODES.RTN_OK)
				return 1;

			return 0;
		}

		/// <summary>
		/// Receives string messages fired as COM events by SetACL
		/// </summary>
		static void MessageHandler(string message)
		{
			// For demo purposes, just print the message
			Console.WriteLine(message);
		}
	}
}
