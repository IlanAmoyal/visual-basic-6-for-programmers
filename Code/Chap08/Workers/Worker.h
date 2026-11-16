////////////////////////////////////////////////////////////////////////////////
// 
// Filename: Worker.h
//
// Purpose: Demonstrate the Procedural Programming Paradigm
//
// Copyrights (C) 2001 by Ilan Amoyal, All Rights Reserved
//
////////////////////////////////////////////////////////////////////////////////

#ifndef WORKER_H
#define WORKER_H

#include <stddef.h>
#include <tchar.h>

#define SIZE_NAME		20			// Max name string-size
#define SIZE_ADDRESS	100			// Max address string-size
#define SIZE_PHONE		20			// Max phone string-size

typedef struct tagWorker
{
	long ID;
	TCHAR Name[SIZE_NAME];
	TCHAR Address[SIZE_ADDRESS];
	TCHAR Phone[SIZE_PHONE];
} Worker;

void SetWorkerData(Worker *pWorker, 
				   long id, 
				   _TCHAR *name, 
				   _TCHAR *address,
				   _TCHAR *phone);

void PrintWorkerData(FILE *out, Worker *pWorker, size_t nWorkers);

#endif // WOEKER_H