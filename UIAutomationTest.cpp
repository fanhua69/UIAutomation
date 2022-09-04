// UIAutomationTest.cpp : Defines the entry point for the application.
//
#include <windows.h>
#include "framework.h"
#include "UIAutomation.h"
#include "UIAutomationTest.h"
#include "UIAutomationClient.h"
#include "UIAutomationCore.h"
#include "UIAutomationCoreApi.h"


#define MAX_LOADSTRING 100

// Global Variables:
HINSTANCE hInst;                                // current instance
WCHAR szTitle[MAX_LOADSTRING];                  // The title bar text
WCHAR szWindowClass[MAX_LOADSTRING];            // the main window class name

// Forward declarations of functions included in this code module:
ATOM                MyRegisterClass(HINSTANCE hInstance);
BOOL                InitInstance(HINSTANCE, int);
LRESULT CALLBACK    WndProc(HWND, UINT, WPARAM, LPARAM);
INT_PTR CALLBACK    About(HWND, UINT, WPARAM, LPARAM);
IUIAutomationElement* GetTopLevelWindowByName(LPWSTR windowName);
IUIAutomationElement* GetTopLevelWindowByID(DWORD processID);
IUIAutomationElement* GetChildWindowByName(IUIAutomationElement* pCurrent, LPWSTR windowName);
IUIAutomationElement* GetChildWindowByNameAndType(IUIAutomationElement* pCurrent, LPWSTR  windowName, long controlID);

#include <uiautomation.h>

// CoInitialize must be called before calling this function, and the  
// caller must release the returned pointer when finished with it.
// 
HRESULT InitializeUIAutomation(IUIAutomation** ppAutomation)
{
    return CoCreateInstance(CLSID_CUIAutomation, NULL,
        CLSCTX_INPROC_SERVER, IID_IUIAutomation,
        reinterpret_cast<void**>(ppAutomation));
}

int APIENTRY wWinMain(_In_ HINSTANCE hInstance,
                     _In_opt_ HINSTANCE hPrevInstance,
                     _In_ LPWSTR    lpCmdLine,
                     _In_ int       nCmdShow)
{
    UNREFERENCED_PARAMETER(hPrevInstance);
    UNREFERENCED_PARAMETER(lpCmdLine);

    // TODO: Place code here.

    // Initialize global strings
    LoadStringW(hInstance, IDS_APP_TITLE, szTitle, MAX_LOADSTRING);
    LoadStringW(hInstance, IDC_UIAUTOMATIONTEST, szWindowClass, MAX_LOADSTRING);
    MyRegisterClass(hInstance);

    // Perform application initialization:
    if (!InitInstance (hInstance, nCmdShow))
    {
        return FALSE;
    }

    STARTUPINFO si;
    PROCESS_INFORMATION pi;
    ZeroMemory(&si, sizeof(si));
    si.cb = sizeof(si);
    ZeroMemory(&pi, sizeof(pi));

    LPTSTR szCmdline = _tcsdup(TEXT("C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE"));
    /*CreateProcess(NULL, szCmdline, NULL,           // Process handle not inheritable
        NULL,           // Thread handle not inheritable
        FALSE,          // Set handle inheritance to FALSE
        0,              // No creation flags
        NULL,           // Use parent's environment block
        NULL,           // Use parent's starting directory 
        &si,            // Pointer to STARTUPINFO structure
        &pi);*/

    WCHAR outlookWindowName[250] = L"Inbox -fanhua69@gmail.com - Outlook";
    IUIAutomationElement* pOutlook = GetTopLevelWindowByID(pi.dwProcessId); // outlookWindowName);
    UIA_HWND phwnd;
    HRESULT hr = pOutlook->get_CurrentNativeWindowHandle(&phwnd);


    if (pOutlook)
    {
        WCHAR newEmail[250] = L"New Email";
        IUIAutomationElement* pNewEmail = GetChildWindowByName(pOutlook, newEmail);
        if (pNewEmail)
        {
            IUIAutomationInvokePattern* pInvokePattern;
            if (SUCCEEDED(pNewEmail->GetCurrentPatternAs(UIA_InvokePatternId,
                __uuidof(IUIAutomationInvokePattern),
                (void**)&pInvokePattern)))
            {
                HRESULT hr = pInvokePattern->Invoke();
            }
        }

        WCHAR new_email_name[250] = L"Untitled - Message (HTML) ";
        IUIAutomationElement* pNewEmailWindow = GetTopLevelWindowByName(new_email_name);
        if(pNewEmailWindow)
        {
            WCHAR to[250] = L"To";
            IUIAutomationElement* pTo = GetChildWindowByNameAndType(pNewEmailWindow, to, UIA_EditControlTypeId);
            if (pTo)
            {
                IUIAutomationValuePattern* pValuePattern = nullptr;
                if (SUCCEEDED(pTo->GetCurrentPatternAs(UIA_ValuePatternId,
                    __uuidof(IUIAutomationValuePattern),
                    (void**)&pValuePattern)))
                {
                    //WCHAR toText[250] = L"gaurang.patel@automationanywhere.com";
                    WCHAR toText[250] = L"fanhua69@gmail.com";
                    HRESULT hr = pValuePattern->SetValue(toText);
                }
            }

            WCHAR subject[250] = L"Subject";
            IUIAutomationElement* pSubject = GetChildWindowByNameAndType(pNewEmailWindow, subject, UIA_EditControlTypeId);
            if (pSubject)
            {
                IUIAutomationValuePattern* pValuePattern = nullptr;
                if (SUCCEEDED(pSubject->GetCurrentPatternAs(UIA_ValuePatternId,
                    __uuidof(IUIAutomationValuePattern),
                    (void**)&pValuePattern)))
                {
                    WCHAR subjectText[250] = L"COM API Test";
                    HRESULT hr = pValuePattern->SetValue(subjectText);
                }
            }

            WCHAR mailBody[250] = L"Page 1 content";
            IUIAutomationElement* pBody = GetChildWindowByNameAndType(pNewEmailWindow, mailBody, UIA_EditControlTypeId);
            if (pBody)
            {
                IUIAutomationValuePattern* pValuePattern = nullptr;
                HRESULT hr = pBody->GetCurrentPatternAs(UIA_ValuePatternId,__uuidof(IUIAutomationValuePattern),(void**)&pValuePattern);
                WCHAR mailBodyText[250] = L"Hello, AutomationEverywhere!";
                if (hr == S_OK && pValuePattern)
                {
                    HRESULT hr = pValuePattern->SetValue(mailBodyText);
                }
                else
                {
                    HRESULT hr = pBody->SetFocus();
                    if (hr == S_OK)
                    {
                        HWND h = (HWND)phwnd;
                        SendMessage(h, WM_SETTEXT, 0, LPARAM(mailBodyText));
                    }
                }
            }

            WCHAR send[250] = L"Send";
            IUIAutomationElement* pSend = GetChildWindowByNameAndType(pNewEmailWindow, send, UIA_ButtonControlTypeId);

            if (pSend)
            {
                IUIAutomationInvokePattern* pInvokePattern;
                if (SUCCEEDED(pSend->GetCurrentPatternAs(UIA_InvokePatternId,
                    __uuidof(IUIAutomationInvokePattern),
                    (void**)&pInvokePattern)))
                {
                    HRESULT hr = pInvokePattern->Invoke();
                }
            }
        }
    }

    HACCEL hAccelTable = LoadAccelerators(hInstance, MAKEINTRESOURCE(IDC_UIAUTOMATIONTEST));
    MSG msg;

    // Main message loop:
    while (GetMessage(&msg, nullptr, 0, 0))
    {
        if (!TranslateAccelerator(msg.hwnd, hAccelTable, &msg))
        {
            TranslateMessage(&msg);
            DispatchMessage(&msg);
        }
    }

    return (int) msg.wParam;
}

IUIAutomationElement* GetTopLevelWindowByName(LPWSTR windowName)
{
    if (windowName == NULL)
    {
        return NULL;
    }

    VARIANT varProp;
    varProp.vt = VT_BSTR;
    varProp.bstrVal = SysAllocString(windowName);
    if (varProp.bstrVal == NULL)
    {
        return NULL;
    }

    IUIAutomationElement* pRoot = NULL;
    IUIAutomationElement* pFound = NULL;
    IUIAutomationCondition* pCondition = nullptr;

    // Get the desktop element. 
    IUIAutomation* pAutomation = nullptr;
    HRESULT hr = InitializeUIAutomation(&pAutomation);
    if (FAILED(hr) || pAutomation == nullptr)
        goto cleanup;

    hr = pAutomation->GetRootElement(&pRoot);
    if (FAILED(hr) || pRoot == NULL)
        goto cleanup;

    // Get a top-level element by name, such as "Program Manager"
    hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varProp, &pCondition);
    if (FAILED(hr))
        goto cleanup;

    pRoot->FindFirst(TreeScope_Subtree /*TreeScope_Children*/, pCondition, &pFound);

cleanup:
    if (pRoot != NULL)
        pRoot->Release();

    if (pCondition != NULL)
        pCondition->Release();

    VariantClear(&varProp);
    return pFound;
}



IUIAutomationElement* GetTopLevelWindowByID(DWORD processID)
{
    VARIANT varProp;
    varProp.vt = VT_UINT;
    varProp.uintVal = processID;

    IUIAutomationElement* pRoot = NULL;
    IUIAutomationElement* pFound = NULL;
    IUIAutomationCondition* pCondition = nullptr;

    // Get the desktop element. 
    IUIAutomation* pAutomation = nullptr;
    HRESULT hr = InitializeUIAutomation(&pAutomation);
    if (FAILED(hr) || pAutomation == nullptr)
        goto cleanup;

    hr = pAutomation->GetRootElement(&pRoot);
    if (FAILED(hr) || pRoot == NULL)
        goto cleanup;

    // Get a top-level element by name, such as "Program Manager"
    hr = pAutomation->CreatePropertyCondition(UIA_ProcessIdPropertyId, varProp, &pCondition);
    if (FAILED(hr))
        goto cleanup;

    pRoot->FindFirst(TreeScope_Subtree /*TreeScope_Children*/, pCondition, &pFound);

cleanup:
    if (pRoot != NULL)
        pRoot->Release();

    if (pCondition != NULL)
        pCondition->Release();

    VariantClear(&varProp);
    return pFound;
}



IUIAutomationElement* GetChildWindowByName(IUIAutomationElement* pCurrent, LPWSTR windowName)
{
    if (windowName == NULL)
    {
        return NULL;
    }

    VARIANT varProp;
    varProp.vt = VT_BSTR;
    varProp.bstrVal = SysAllocString(windowName);
    if (varProp.bstrVal == NULL)
    {
        return NULL;
    }

    IUIAutomationElement    * pRoot      = NULL;
    IUIAutomationElement    * pFound     = NULL;
    IUIAutomationCondition  * pCondition = nullptr;

    // Get the desktop element. 
    IUIAutomation* pAutomation = nullptr;
    HRESULT hr = InitializeUIAutomation(&pAutomation);
    if (FAILED(hr) || pAutomation == nullptr)
        goto cleanup;

        // Get a top-level element by name, such as "Program Manager"
    hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varProp, &pCondition);
    if (FAILED(hr))
        goto cleanup;

    pCurrent->FindFirst(TreeScope_Subtree /*TreeScope_Children*/, pCondition, &pFound);

cleanup:
    if (pRoot != NULL)
        pRoot->Release();

    if (pCondition != NULL)
        pCondition->Release();

    VariantClear(&varProp);
    return pFound;
}


IUIAutomationElement* GetChildWindowByNameAndType(  IUIAutomationElement* pCurrent, 
                                                    LPWSTR  windowName, 
                                                    long    controlID)
{
    if (windowName == NULL)
    {
        return NULL;
    }

    VARIANT varPropName;
    varPropName.vt = VT_BSTR;
    varPropName.bstrVal = SysAllocString(windowName);
    if (varPropName.bstrVal == NULL)
    {
        return NULL;
    }

    VARIANT varPropType;
    varPropType.vt = VT_INT;
    varPropType.intVal = controlID; // UIA_EditControlTypeId;

    IUIAutomationElement* pRoot = NULL;
    IUIAutomationElement* pFound = NULL;
    IUIAutomationCondition* pConditionName = nullptr;
    IUIAutomationCondition* pConditionType = nullptr;
    IUIAutomationCondition* pConditionBoth = nullptr;


    // Get the desktop element. 
    IUIAutomation* pAutomation = nullptr;
    HRESULT hr = InitializeUIAutomation(&pAutomation);
    if (FAILED(hr) || pAutomation == nullptr)
        goto cleanup;

    // Get a top-level element by name, such as "Program Manager"
    hr = pAutomation->CreatePropertyCondition(UIA_NamePropertyId, varPropName, &pConditionName);
    if (FAILED(hr))
        goto cleanup;

    hr = pAutomation->CreatePropertyCondition(UIA_ControlTypePropertyId, varPropType, &pConditionType);
    if (FAILED(hr))
        goto cleanup;

    hr = pAutomation->CreateAndCondition(pConditionName, pConditionType, &pConditionBoth);
    if (FAILED(hr))
        goto cleanup;

    pCurrent->FindFirst(TreeScope_Subtree /*TreeScope_Children*/, pConditionBoth, &pFound);

cleanup:
    if (pRoot != NULL)
        pRoot->Release();

    if (pConditionName != NULL)
        pConditionName->Release();

    if (pConditionType != NULL)
        pConditionType->Release();

    if (pConditionBoth != NULL)
        pConditionBoth->Release();

    VariantClear(&varPropName);
    VariantClear(&varPropType);
    return pFound;
}

//
//  FUNCTION: MyRegisterClass()
//
//  PURPOSE: Registers the window class.
//
ATOM MyRegisterClass(HINSTANCE hInstance)
{
    WNDCLASSEXW wcex;

    wcex.cbSize = sizeof(WNDCLASSEX);

    wcex.style          = CS_HREDRAW | CS_VREDRAW;
    wcex.lpfnWndProc    = WndProc;
    wcex.cbClsExtra     = 0;
    wcex.cbWndExtra     = 0;
    wcex.hInstance      = hInstance;
    wcex.hIcon          = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_UIAUTOMATIONTEST));
    wcex.hCursor        = LoadCursor(nullptr, IDC_ARROW);
    wcex.hbrBackground  = (HBRUSH)(COLOR_WINDOW+1);
    wcex.lpszMenuName   = MAKEINTRESOURCEW(IDC_UIAUTOMATIONTEST);
    wcex.lpszClassName  = szWindowClass;
    wcex.hIconSm        = LoadIcon(wcex.hInstance, MAKEINTRESOURCE(IDI_SMALL));

    return RegisterClassExW(&wcex);
}

//
//   FUNCTION: InitInstance(HINSTANCE, int)
//
//   PURPOSE: Saves instance handle and creates main window
//
//   COMMENTS:
//
//        In this function, we save the instance handle in a global variable and
//        create and display the main program window.
//
BOOL InitInstance(HINSTANCE hInstance, int nCmdShow)
{
   hInst = hInstance; // Store instance handle in our global variable

   HWND hWnd = CreateWindowW(szWindowClass, szTitle, WS_OVERLAPPEDWINDOW,
      CW_USEDEFAULT, 0, CW_USEDEFAULT, 0, nullptr, nullptr, hInstance, nullptr);

   if (!hWnd)
   {
      return FALSE;
   }

   ShowWindow(hWnd, nCmdShow);
   UpdateWindow(hWnd);

   return TRUE;
}

//
//  FUNCTION: WndProc(HWND, UINT, WPARAM, LPARAM)
//
//  PURPOSE: Processes messages for the main window.
//
//  WM_COMMAND  - process the application menu
//  WM_PAINT    - Paint the main window
//  WM_DESTROY  - post a quit message and return
//
//
LRESULT CALLBACK WndProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
    switch (message)
    {
    case WM_COMMAND:
        {
            int wmId = LOWORD(wParam);
            // Parse the menu selections:
            switch (wmId)
            {
            case IDM_ABOUT:
                DialogBox(hInst, MAKEINTRESOURCE(IDD_ABOUTBOX), hWnd, About);
                break;
            case IDM_EXIT:
                DestroyWindow(hWnd);
                break;
            default:
                return DefWindowProc(hWnd, message, wParam, lParam);
            }
        }
        break;
    case WM_PAINT:
        {
            PAINTSTRUCT ps;
            HDC hdc = BeginPaint(hWnd, &ps);
            // TODO: Add any drawing code that uses hdc here...
            EndPaint(hWnd, &ps);
        }
        break;
    case WM_DESTROY:
        PostQuitMessage(0);
        break;
    default:
        return DefWindowProc(hWnd, message, wParam, lParam);
    }
    return 0;
}

// Message handler for about box.
INT_PTR CALLBACK About(HWND hDlg, UINT message, WPARAM wParam, LPARAM lParam)
{
    UNREFERENCED_PARAMETER(lParam);
    switch (message)
    {
    case WM_INITDIALOG:
        return (INT_PTR)TRUE;

    case WM_COMMAND:
        if (LOWORD(wParam) == IDOK || LOWORD(wParam) == IDCANCEL)
        {
            EndDialog(hDlg, LOWORD(wParam));
            return (INT_PTR)TRUE;
        }
        break;
    }
    return (INT_PTR)FALSE;
}
