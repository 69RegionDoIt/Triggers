#define _WIN32_DCOM
#include <windows.h>
#include <iostream>
#include <stdio.h>
#include <comdef.h>
#include <string.h>
#include <wincred.h>
#include <taskschd.h>
#pragma comment(lib, "taskschd.lib")
#pragma comment(lib, "comsupp.lib")
#pragma comment(lib, "credui.lib")
using namespace std;

int createTask(LPCWSTR taskName, BSTR taskPath, BSTR query, BSTR args, BOOL putValueQueries, BSTR valueName, BSTR valueValue) {
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		printf("\nCoInitializeEx failed: %x", hr);
		return 1;
	}

	hr = CoInitializeSecurity(NULL,-1,NULL,NULL,RPC_C_AUTHN_LEVEL_PKT_PRIVACY,RPC_C_IMP_LEVEL_IMPERSONATE, NULL, 0, NULL);
	if (FAILED(hr)) {
		printf("\nCoInitializeSecurity failed: %x", hr);
		CoUninitialize();
		return 1;
	}

	ITaskService *pService = NULL;
	hr = CoCreateInstance(CLSID_TaskScheduler, NULL, CLSCTX_INPROC_SERVER, IID_ITaskService, (void**)&pService);
	if (FAILED(hr)) {
		printf("Failed to CoCreate an instance of the TaskService class: %x", hr);
		CoUninitialize();
		return 1;
	}

	hr = pService->Connect(_variant_t(), _variant_t(), _variant_t(), _variant_t());
	if (FAILED(hr)) {
		printf("ITaskService::Connect failed: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}

	ITaskFolder *pRootFolder = NULL;
	hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
	
	if (FAILED(hr)) {
		printf("Cannot get Root Folder pointer: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}
	
	pRootFolder->DeleteTask(_bstr_t(taskName), 0);
	ITaskDefinition *pTask = NULL;
	hr = pService->NewTask(0, &pTask);
	pService->Release();
	if (FAILED(hr)) {
		printf("Failed to create an instance of the task: %x", hr);
		pRootFolder->Release();
		CoUninitialize();
		return 1;
	}
	
	IRegistrationInfo *pRegInfo = NULL;
	hr = pTask->get_RegistrationInfo(&pRegInfo);
	if (FAILED(hr)) {
		printf("\nCannot get identification pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	
	hr = pRegInfo->put_Author(L"Nikita Gulin");
	pRegInfo->Release();
	if (FAILED(hr)) {
		printf("\nCannot put identification info: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	ITaskSettings *pSettings = NULL;
	hr = pTask->get_Settings(&pSettings);
	if (FAILED(hr)) {
		printf("\nCannot get settings pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	
	hr = pSettings->put_StartWhenAvailable(VARIANT_TRUE);
	pSettings->Release();
	if (FAILED(hr)) {
		printf("\nCannot put setting info: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	
	ITriggerCollection *pTriggerCollection = NULL;
	hr = pTask->get_Triggers(&pTriggerCollection);
	if (FAILED(hr)) {
		printf("\nCannot get trigger collection: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	
	ITrigger *pTrigger = NULL;
	hr = pTriggerCollection->Create(TASK_TRIGGER_EVENT, &pTrigger);
	pTriggerCollection->Release();
	if (FAILED(hr)) {
		printf("\nCannot create the trigger: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	
	IEventTrigger *pEventTrigger = NULL;
	hr = pTrigger->QueryInterface(IID_IEventTrigger, (void**)&pEventTrigger);
	pTrigger->Release();
	if (FAILED(hr)) {
		printf("\nQueryInterface call on IEventTrigger failed: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}

	hr = pEventTrigger->put_Id(_bstr_t(L"Trigger1"));
	if (FAILED(hr))
		printf("\nCannot put the trigger ID: %x", hr);
	hr = pEventTrigger->put_Subscription(query);
	if (FAILED(hr)) {
		printf("\nCannot put the event query: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	if (putValueQueries == TRUE) {
		ITaskNamedValueCollection * pValueQueries = NULL;
		pEventTrigger->get_ValueQueries(&pValueQueries);
		pValueQueries->Create(valueName, valueValue, NULL);
		hr = pEventTrigger->put_ValueQueries(pValueQueries);
		if (FAILED(hr)) {
			printf("\nCannot put value queries: %x", hr);
			pRootFolder->Release();
			pTask->Release();
			CoUninitialize();
			return 1;
		}
	}
	pEventTrigger->Release();
	IActionCollection *pActionCollection = NULL;
	hr = pTask->get_Actions(&pActionCollection);
	if (FAILED(hr)) {
		printf("\nCannot get action collection pointer: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	IAction *pAction = NULL;
	hr = pActionCollection->Create(TASK_ACTION_EXEC, &pAction);
	pActionCollection->Release();
	if (FAILED(hr)) {
		printf("\nCannot create an exec action: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	IExecAction *pExecAction = NULL;
	hr = pAction->QueryInterface(IID_IExecAction, (void**)&pExecAction);
	pAction->Release();
	if (FAILED(hr)) {
		printf("\nQueryInterface call failed for IEmailAction: %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	hr = pExecAction->put_Path(taskPath);
	if (FAILED(hr)) {
		printf("\nCannot put path information: %x", hr);
		pRootFolder->Release();
		pExecAction->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	hr = pExecAction->put_Arguments(args);
	if (FAILED(hr)) {
		printf("\nCannot put arguments information: %x", hr);
		pRootFolder->Release();
		pExecAction->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	pExecAction->Release();
	IRegisteredTask *pRegisteredTask = NULL;
	hr = pRootFolder->RegisterTaskDefinition(
		_bstr_t(taskName),
		pTask,
		TASK_CREATE_OR_UPDATE,
		_variant_t(_bstr_t(L"")),
		_variant_t(_bstr_t(L"")),
		TASK_LOGON_INTERACTIVE_TOKEN,
		_variant_t(L""),
		&pRegisteredTask);
	if (FAILED(hr)) {
		printf("\nError saving the Task : %x", hr);
		pRootFolder->Release();
		pTask->Release();
		CoUninitialize();
		return 1;
	}
	wcout << "Success! Task successfully registered." << endl;
	pRootFolder->Release();
	pTask->Release();
	pRegisteredTask->Release();
	CoUninitialize();
	return 0;
}

int printTasksRecursive(ITaskFolder * taskFolder, int depth) {
	IRegisteredTaskCollection* pTaskCollection = NULL;
	HRESULT hr = taskFolder->GetTasks(NULL, &pTaskCollection);
	ITaskFolderCollection* pFolderCollection = NULL;
	hr = taskFolder->GetFolders(NULL, &pFolderCollection);
	BSTR folderName;
	taskFolder->get_Name(&folderName);
	taskFolder->Release();
	if (FAILED(hr)) {
		wcout << "Cannot get the registered tasks or folders.: " <<  hr << endl;
		CoUninitialize();
		return 1;
	}
	LONG numTasks = 0;
	hr = pTaskCollection->get_Count(&numTasks);
	LONG numFolders = 0;
	hr = pFolderCollection->get_Count(&numFolders);

	SysFreeString(folderName);
	if (numTasks > 0) {
		TASK_STATE taskState;
		for (LONG i = 0; i < numTasks; i++) {
			IRegisteredTask* pRegisteredTask = NULL;
			hr = pTaskCollection->get_Item(_variant_t(i + 1), &pRegisteredTask);
			if (SUCCEEDED(hr)) {
				BSTR taskName = NULL;
				hr = pRegisteredTask->get_Name(&taskName);
				if (SUCCEEDED(hr)) {
					hr = pRegisteredTask->get_State(&taskState);
					if (SUCCEEDED(hr)) {
						if(! taskState)
							wcout << taskName << " Unknown" << endl;
						else if(taskState == 1)
							wcout << taskName << " Disabled" << endl;
						else if(taskState == 2)
							wcout << taskName << " Queued" << endl;
						else if(taskState == 3)
							wcout << taskName << " Ready" << endl;
						else
							wcout << taskName << " Running" << endl;
						SysFreeString(taskName);
					}
					else {
						wcout << "Task Name:" << taskName <<"state: couldn't get state" << endl;
						SysFreeString(taskName);
					}
				}
				else {
					wcout << "Cannot get the registered task name: " <<  hr << endl;
				}
				pRegisteredTask->Release();
			}
			else {
				wcout << "Cannot get the registered task item at index = " << i + 1 << " : " << hr;
			}
		}
	}
	pTaskCollection->Release();
	for (LONG i = 0; i < numFolders; i++) {
		ITaskFolder* subFolder = NULL;
		pFolderCollection->get_Item(_variant_t(i + 1), &subFolder);
		printTasksRecursive(subFolder, depth + 1);
	}
	pFolderCollection->Release();
	return 0;
}

int Show_tasks() {
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		printf("\nCoInitializeEx failed: %x", hr);
		return 1;
	}
	hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		0,
		NULL);
	if (FAILED(hr)) {
		printf("\nCoInitializeSecurity failed: %x", hr);
		CoUninitialize();
		return 1;
	}
	ITaskService *pService = NULL;
	hr = CoCreateInstance(CLSID_TaskScheduler,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ITaskService,
		(void**)&pService);
	if (FAILED(hr)) {
		printf("Failed to CoCreate an instance of the TaskService class: %x", hr);
		CoUninitialize();
		return 1;
	}
	hr = pService->Connect(_variant_t(), _variant_t(),
		_variant_t(), _variant_t());
	if (FAILED(hr)) {
		printf("ITaskService::Connect failed: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}
	ITaskFolder *pRootFolder = NULL;
	hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
	pService->Release();
	if (FAILED(hr)) {
		printf("Cannot get Root Folder pointer: %x", hr);
		CoUninitialize();
		return 1;
	}
	printTasksRecursive(pRootFolder, 1);
	CoUninitialize();
	printf("\n");
	return 0;
}

int deleteTask(LPCWSTR taskName) {
	HRESULT hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hr)) {
		printf("\nCoInitializeEx failed: %x", hr);
		return 1;
	}
	hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_PKT_PRIVACY,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		NULL,
		0,
		NULL);
	if (FAILED(hr)) {
		printf("\nCoInitializeSecurity failed: %x", hr);
		CoUninitialize();
		return 1;
	}
	ITaskService *pService = NULL;
	hr = CoCreateInstance(CLSID_TaskScheduler,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_ITaskService,
		(void**)&pService);
	if (FAILED(hr)) {
		printf("Failed to CoCreate an instance of the TaskService class: %x", hr);
		CoUninitialize();
		return 1;
	}
	hr = pService->Connect(_variant_t(), _variant_t(),
		_variant_t(), _variant_t());
	if (FAILED(hr)) {
		printf("ITaskService::Connect failed: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}
	ITaskFolder *pRootFolder = NULL;
	hr = pService->GetFolder(_bstr_t(L"\\"), &pRootFolder);
	if (FAILED(hr)) {
		printf("Cannot get Root Folder pointer: %x", hr);
		pService->Release();
		CoUninitialize();
		return 1;
	}
	pRootFolder->DeleteTask(_bstr_t(taskName), 0);
	printf("Task was deleted.\n");
	pRootFolder->Release();
	CoUninitialize();
	return 0;
}

int main(int argc, char ** argv)
{
	if (argc < 1)
		return 0;
	if (!strcmp(argv[1], "Show_tasks")) {
		Show_tasks();
	}
	else if (!strcmp(argv[1], "Event_firewall"))
	{
		createTask(L"firewallWarning", L"C:\\Users\\acer\\Documents\\Visual Studio 2015\\Projects\\BsitLab3\\Debug\\BsitLab3.exe",
			L"<QueryList>\
			<Query Id=\"0\" Path=\"Microsoft-Windows-Windows Firewall With Advanced Security/Firewall\">\
			<Select Path=\"Microsoft-Windows-Windows Firewall With Advanced Security/Firewall\">*</Select>\
			</Query>\
			</QueryList>",
			L"firewall",
			FALSE, L"", L"");
	}
	else if (!strcmp(argv[1], "Event_defender"))
	{
		createTask(L"defenderWarning", L"C:\\Users\\acer\\Documents\\Visual Studio 2015\\Projects\\BsitLab3\\Debug\\BsitLab3.exe",
			L"<QueryList>\
			<Query Id=\"0\" Path=\"Microsoft-Windows-Windows Defender/Operational\">\
			<Select Path=\"Microsoft-Windows-Windows Defender/Operational\">*[System[Provider[@Name='Microsoft-Windows-Windows Defender'] and (EventID = 5007)]]</Select>\
			<Select Path=\"Microsoft-Windows-Windows Defender/WHC\">*[System[Provider[@Name='Microsoft-Windows-Windows Defender'] and (EventID = 5007)]]</Select>\
			</Query>\
			</QueryList>",
			L"defender",
			FALSE, L"", L"");
	}
	else if (strcmp(argv[1], "Event_ping") == 0) {
		createTask(L"WFPWarning", L"C:\\Users\\acer\\Documents\\Visual Studio 2015\\Projects\\BsitLab3\\Debug\\BsitLab3.exe",
			L"<QueryList>\
			<Query Id=\"0\" Path=\"Security\">\
			<Select Path=\"Security\">*[System[(EventID = 5152)]]</Select>\
			</Query>\
			</QueryList>",
			L"ping $(srcIp)",
			TRUE,
			L"srcIp",
			L"Event/EventData/Data[@Name='SourceAddress']"
		);
	}
	else if (!strcmp(argv[1], "Remove_defender"))
	{
		deleteTask(L"defenderWarning");
	}
	else if (!strcmp(argv[1], "prostoi"))
	{
		wcout << "Prostoi" << endl;
		system("pause");
	}
	else if (!strcmp(argv[1], "Remove_firewall"))
	{
		deleteTask(L"firewallWarning");
	}
	else if (!strcmp(argv[1], "Remove_ping")) {
		deleteTask(L"WFPWarning");
	}
	else if (strcmp(argv[1], "firewall") == 0) {
		wcout << "Windows Firewall settings were changed!" << endl;
		system("pause");
	}
	else if (strcmp(argv[1], "defender") == 0) {
		wcout << "Windows Defender settings were changed!" << endl;
		system("pause");
	}
	else if (strcmp(argv[1], "ping") == 0) {
		char dropped_ip[17] = "Some ip address";
			if (!strcmp(dropped_ip, argv[2])) {
				wcout << "Ping from " << argv[2] << " was dropped." << endl;
				system("pause");
			}
	}
	else
	{
		wcout << "Show_tasks: show all tasks and theirs state." << endl;
		wcout << "Event_defender: activate task, which show message about windows defender changes." << endl;
		wcout << "Event_firewall: activate task, which show message about firewall changes." << endl;
		wcout << "Event_ping: activate task, which show message about dropped pings." << endl;
		wcout << "Remove_defender: stopped task, which show message about windows defender changes." << endl;
		wcout << "Remove_firewall: stopped task, which show message about firewall pings." << endl;
		wcout << "Remove_ping: stopped task, which show message about dropped pings." << endl;
	}
}