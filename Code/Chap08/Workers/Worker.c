////////////////////////////////////////////////////////////////////////////////
// 
// Filename: Worker.c
//
// Purpose: Demonstrate the Procedural Programming Paradigm
//
// Copyrights (C) 2001 by Ilan Amoyal, All Rights Reserved
//
////////////////////////////////////////////////////////////////////////////////

#include <stdio.h>
#include "Worker.h"

void SetWorkerData(Worker *pWorker, long id, _TCHAR *name, 
				   _TCHAR *address, _TCHAR *phone)
{
	pWorker->ID = id;
	_tcsncpy(pWorker->Name, name, SIZE_NAME - 1);
	_tcsncpy(pWorker->Address, address, SIZE_ADDRESS - 1);
	_tcsncpy(pWorker->Phone, phone, SIZE_PHONE - 1);
}

void PrintWorkerData(FILE *out, Worker *pWorker, 
					 size_t nWorkers)
{
	size_t i;
	for(i = 0; i < nWorkers; i++)
	{
		_ftprintf(out, 
				  _T("ID: %ld\n"
					 "Name: %s\n"
					 "Address: %s\n"
					 "Phone: %s\n\n"), 
				  (pWorker + i)->ID,
				  (pWorker + i)->Name,
				  (pWorker + i)->Address,
				  (pWorker + i)->Phone);
	}
}


