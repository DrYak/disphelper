/* This file contains sample code that demonstrates use of the DispHelper COM helper library.
 * DispHelper allows you to call COM objects with an extremely simple printf style syntax.
 * DispHelper can be used from C++ or even plain C. It works with most Windows compilers
 * including Dev-CPP, Visual C++ and LCC-WIN32. Including DispHelper in your project
 * couldn't be simpler as it is available in a compacted single file version.
 *
 * Included with DispHelper are over 20 samples that demonstrate using COM objects
 * including ADO, CDO, Outlook, Eudora, Excel, Word, Internet Explorer, MSHTML,
 * PocketSoap, Word Perfect, MS Agent, SAPI, MSXML, WIA, dexplorer and WMI.
 *
 * DispHelper is free open source software provided under the BSD license.
 *
 * Find out more and download DispHelper at:
 * http://sourceforge.net/projects/disphelper/
 * http://disphelper.sourceforge.net/
 *
 * You can browse the source code and samples, including a C version of this sample, online at:
 * http://cvs.sourceforge.net/viewcvs.py/disphelper/disphelper/
 */


/* --
word.cpp:
  Demonstrates outputting formatted text to a Word document and getting user feed
back with the help of the office assistant. Demonstrates using Word as a spell
checker.
 -- */


#include "disphelper.h"
#include <iostream>
#include <string>
using namespace std;


/* **************************************************************************
 * Word Sample 1:
 *   Demonstrates outputting text to a word document and getting feedback from
 * the user with the office assistant.
 *
 ============================================================================ */
int WordSample1(void)
{
	CDhInitialize init;    /* Initialize DispHelper for this thread. */
	CDispPtr wdApp, wdDoc, wdBalloon;
	UINT iChoice, i=0;

	/* Turn on DispHelper exceptions */
	dhToggleExceptions(TRUE);

	try
	{
		/* Create an instance of 'Word.Application' on the local machine */
		dhCheck( dhCreateObject(L"Word.Application", NULL, &wdApp) );

		/* wdApp.Visible = True */
		dhPutValue(wdApp, L".Visible = %b", TRUE);

		/* Set wdDoc = wdApp.Documents.Add */
		/* '%o' means we wish to receive an IDispatch object in wdDoc */
		dhCheck( dhGetValue(L"%o", &wdDoc, wdApp, L".Documents.Add") );

		/* Turn off those red squigly lines */
		dhPutValue(wdApp, L".ActiveDocument.ShowSpellingErrors = %b", FALSE);

		/* Set the font properties for the document */
		WITH(wdFont, wdDoc, L".Range.Font")
		{
			dhPutValue(wdFont, L".Size = %d", 20);
			dhPutValue(wdFont, L".SmallCaps = %b", TRUE);
			dhPutValue(wdFont, L".Name = %S", L"Comic Sans MS");
			dhPutValue(wdFont, L".Color = %d", RGB(33,99,0) );   /* A nice green */

		} END_WITH(wdFont); /* Always remember the END_WITH! */
	
		/* Some different ways to add text */

		/* Call wdApp.Selection.TypeText with a wide('%S') string */
		dhCallMethod(wdApp, L".Selection.TypeText(%S)", L"DispHelper Word Sample!\n\n");

		/* Call wdApp.Selection.TypeText with an ansi('%s') string */
		dhCallMethod(wdApp, L".Selection.TypeText(%s)", "DispHelper allows the C/C++ programmer ");

		/* Call wdDoc.Range.InsertAfter with a LPTSTR('%T') */
		dhCallMethod(wdDoc, L".Range.InsertAfter(%T)", TEXT("to make easy use of COM objects.\n"));

		/* Take advantage of automatic conversion to insert the time as a time_t('%t') */
		dhCallMethod(wdDoc, L".Range.InsertAfter(%t)", time(NULL));

		/* wdDoc.Shapes.AddTextEffect 15 , "DispHelper COM Helper!" , "Arial Black" , 36 , 0 , 0 , 90.0, 224.95 */
		dhCallMethod(wdDoc, L".Shapes.AddTextEffect(%d,%S,%S,%d,%d,%d,%e,%e)", 15, L"DispHelper COM Helper!", L"Arial Black", 36, 0, 0, 90.0, 224.95);

		/* Now rotate the text effect (the first item in the shapes collection) */
		dhPutValue(wdDoc, L".Shapes(%d).Rotation = %e", 1, -20.5);

		/* Use the office assistant to get input from the user */
	
		/* Set wdBalloon = wdApp.Assistant.NewBalloon */
		dhCheck( dhGetValue(L"%o", &wdBalloon, wdApp, L".Assistant.NewBalloon") );

		/* wdBalloon.BalloonType = 0 'msoBalloonTypeButtons */
		dhPutValue(wdBalloon, L".BalloonType = %d", 0);

		/* wdBalloon.Text = "Which color scheme would you like?" */
		dhPutValue(wdBalloon, L".Text = %S", L"Which color scheme would you like?");

		/* wdBalloon.Labels(1).Text = "Crimson/Indigo" */
		dhPutValue(wdBalloon, L".Labels(%d).Text = %S", 1, L"Crimson/Indigo");

		/* wdBalloon.Labels(2).Text = "Orange/Turquoise" */
		dhPutValue(wdBalloon, L".Labels(%d).Text = %S", 2, L"Orange/Turquoise");

		/* iChoice = wdBalloon.Show */
		dhGetValue(L"%d", &iChoice, wdBalloon, L".Show");

		/* Loop through each word in the document and set its color based on the user's choice */
		FOR_EACH(wdWord, /* IN */ wdDoc, L".Words")
		{
			COLORREF clrf;

			if (iChoice == 1) clrf = (i % 2 ? RGB(0xDC,0x14,0x3C) : RGB(0x4B,0x00,0x82)); /* Alternate crimson/indigo */
			else              clrf = (i % 2 ? RGB(0xFF,0x45,0x00) : RGB(0x40,0xE0,0xD0)); /* Alternate orange/turquoise */

			dhPutValue(wdWord, L".Font.Color = %d", clrf);
			dhPutValue(wdWord, L".Font.Bold = %d", !(i++ % 3)); /* Make every 3rd word bold. */

			Sleep(50); /* Pause for effect. */

		} NEXT_THROW(wdWord); /* Always remember the NEXT! */

		/* wdDoc.WebPagePreview() */
		/* This will look like the web site of a two year old! */
		dhCallMethod(wdDoc, L".WebPagePreview");
	}
	catch (string errstr)
	{
		cerr << "Fatal error details:" << endl << errstr << endl;
	}

	/* Turn off exceptions display for cleanup */
	dhToggleExceptions(FALSE);

	/* wdApp.ActiveDocument.Saved = True */
	/* This prevents the user being prompted when word is closed. */
	dhPutValue(wdApp, L".ActiveDocument.Saved = %b", TRUE);

	/* wdApp.Quit() */
	dhCallMethod(wdApp, L".Quit");

	return 0;
}


/* **************************************************************************
 * Word Sample 2:
 *   Demonstrates using word as a spell checker.
 *
 ============================================================================ */
int WordSample2(void)
{
	CDhInitialize init;
	CDispPtr wdDoc;
	const char * szText = "Thid is a te'st of using Worfd as a spellifng checker. The quicf fox atee the brown kangarloo.";

	dhToggleExceptions(TRUE);

	try
	{
		cout << "Checking spelling of '" << szText << "'..." << endl;

		/* Create a new word document */
		dhCheck( dhCreateObject(L"Word.Document", NULL, &wdDoc) );

		/* Set document text */
		dhCheck( dhPutValue(wdDoc, L".Range.Text = %s", szText) );

		/* Enumerate all the misspelled words */
		FOR_EACH(wdRange, wdDoc, L".SpellingErrors")
		{
			CDhStringA szWord;

			/* Retrieve and print out the misspelled word */
			dhGetValue(L"%s", &szWord, wdRange, L".Text");
			cout << "\nMisspelled Word: " << szWord << endl;

			/* Now retrieve the suggestions */
			FOR_EACH(wdSuggestion, wdRange, L".GetSpellingSuggestions")
			{
				CDhString szSuggestion;
	
				dhGetValue(L"%s", &szSuggestion, wdSuggestion, L".Name");
				cout << "\tSuggestion: " << szSuggestion << endl;

			} NEXT_THROW(wdSuggestion);

		} NEXT_THROW(wdRange);
	}
	catch (string errstr)
	{
		cerr << "Fatal error details:" << endl << errstr << endl;
	}

	/* Close document(0 means wdDoNotSaveChanges) */
	dhCallMethod(wdDoc, L".Close(%d)", 0);

	return 0;
}


/* ============================================================================ */
int main(void)
{
	cout << "Running Word Sample One..." << endl;
	WordSample1();

	cout << "\nPress ENTER to run Word Sample Two..." << endl;
	cin.get();
	WordSample2();

	cout << "\nPress ENTER to exit..." << endl;
	cin.get();

	return 0;
}


/* Word Notes:

- Word seems to use '\r' as its solitary new line character. This will cause
unexpected results when output to the console using printf, puts, etc.

- Opening an existing document and getting its contents(as a LPTSTR).
dhCreateObject(L"Word.Application, NULL, &wdApp);
dhGetValue(L"%o", &wdDoc, wdApp, L".Documents.Open(%s)", "mydoc.doc");
dhGetValue(L"%T", &szText, wdDoc, L".Range.Text");

- Running a macro.
dhCallMethod(wdApp, L".Run(%s, %d, %S)", "'My Document.doc'!ThisModule.ThisProcedure", 1, L"arg2");
The document must be open. The first argument is the macro name. This can
be followed by macro arguments.

- Get a running instance of word.
dhGetObject(NULL, L"Word.Application", &wdApp);

- Enumerate each document.
FOR_EACH(wdDoc, wdApp, L".Documents")
{
	// Do something here...
} NEXT(wdDoc);

*/

