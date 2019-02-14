#include <windows.h>
#include <stdio.h>
#include <ole2.h>

// Add the following libraries to the linker script:
// -lole32
// -luuid
// -loleaut32

BOOL CreateSpeechObject(void** speechDispatch);
void Speak(IDispatch* speechDispatch, OLECHAR* textToSpeak);

int WINAPI WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow) {
	
	IDispatch* speechDispatch = NULL;
	
	OleInitialize(NULL);
	CoInitialize(NULL);
	if(!CreateSpeechObject((void**)&speechDispatch)) return 1;
	Speak(speechDispatch, L"Hello everyone. I am very tired. So I am going to go to sleep. Bye bye.");
	CoUninitialize();
	OleUninitialize();
	system("PAUSE");
	
}

BOOL CreateSpeechObject(void** speechDispatch) {

	CLSID speechClass;
	
	if(CLSIDFromProgID(L"Sapi.SpVoice", &speechClass) == NOERROR) {		
		if(CoCreateInstance(&speechClass, NULL, CLSCTX_ALL, &IID_IDispatch, speechDispatch) == NOERROR) {		
			return TRUE;
		}
	}
	
	return FALSE;
	
}

void Speak(IDispatch* speechDispatch, OLECHAR* textToSpeak) {
	
	DISPID speakID;
	DISPPARAMS dispatchParams;
	OLECHAR** speechProcName;
	VARIANT textVar;
	
	memset(&dispatchParams, 0, sizeof(DISPPARAMS));
	
	OLECHAR* procName = L"Speak";
	speechProcName = &procName;
	
	BSTR textStr = SysAllocString(textToSpeak);
	
	speechDispatch->lpVtbl->GetIDsOfNames(speechDispatch, &IID_NULL, speechProcName, 1, LOCALE_USER_DEFAULT, &speakID);
	
	VariantInit(&textVar);
    textVar.vt = VT_BSTR;
    textVar.bstrVal = textStr;
	
	dispatchParams.cArgs = 1;
	dispatchParams.rgvarg = &textVar;
	
	speechDispatch->lpVtbl->Invoke(speechDispatch, speakID, &IID_NULL, LOCALE_SYSTEM_DEFAULT, DISPATCH_METHOD, &dispatchParams, NULL, NULL, NULL);
	
	VariantClear(&textVar);
	SysFreeString(textStr);
	
}	
