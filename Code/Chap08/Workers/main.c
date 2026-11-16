////////////////////////////////////////////////////////////////////////////////
// 
// Filename: main.c
//
// Purpose: Demonstrate the Procedural Programming Paradigm
//
// Copyrights (C) 2001 by Ilan Amoyal, All Rights Reserved
//
////////////////////////////////////////////////////////////////////////////////

#include <stdio.h>
#include <tchar.h>
#include "Worker.h"

#define SIZE_WORKERS		20			// Define the workers array size

// define menu options
#define OPT_QUIT			0
#define OPT_ADD				1
#define OPT_PRINT			2

int ShowMenu();

int _tmain(int argc, _TCHAR **argv, _TCHAR **envp)
{
	size_t nWorkers = 0;
	int nOption;
	long id;
	_TCHAR name[SIZE_NAME], address[SIZE_ADDRESS], phone[SIZE_PHONE];
	Worker workers[SIZE_WORKERS];

	while((nOption = ShowMenu()) != OPT_QUIT)
	{
		switch(nOption)
		{
		case OPT_ADD:
			if(nWorkers < SIZE_WORKERS)
			{
				_tprintf(_T("Enter worker ID:"));
				_tscanf(_T("%ld"), &id);
				_tprintf(_T("Enter worker name:"));
				_tscanf(_T("%s"), name);
				_tprintf(_T("Enter worker address:"));
				_tscanf(_T("%s"), address);
				_tprintf(_T("Enter worker phone:"));
				_tscanf(_T("%s"), phone);
				SetWorkerData(&workers[nWorkers], id, name, address, phone);
				nWorkers++;
			}
			else
			{
				_tprintf(_T("Error: Can't add another worker. "
						"The array is full\n"));
			} // end if(nWorkers...
			break;
		case OPT_PRINT:
			PrintWorkerData(stdout, workers, nWorkers);
			break;
		default:
			_tprintf(_T("Error: Unknown option. Try again.\n\n"));
			break;
		} // end switch
	} // end of while
	return 0;
}

int ShowMenu()
{
	int nOption;
	_tprintf(_T("\n"
				"Menu:\n"
			    "%d - Quit\n"
			    "%d - Add worker\n"
			    "%d - Print\n"
				"Enter your choice here:"), 
				OPT_QUIT, OPT_ADD, OPT_PRINT);
	_tscanf(_T("%d"), &nOption);
	return nOption;
}