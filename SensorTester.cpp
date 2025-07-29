// SensorTester.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <string>
#include <iomanip>

#include <Windows.h>
#include <comutil.h>
#include <objbase.h>
#include <sensorsapi.h>
#include <sensors.h>
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "ole32.lib")
#pragma comment(lib, "sensorsapi.lib")

using namespace std;
using namespace _com_util;


int main()
{
    HRESULT hr = S_OK;

    ISensorManager* pSensorManager = NULL;
    ISensorCollection* pSensorCollection = NULL;

    hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (SUCCEEDED(hr))
        hr = CoCreateInstance(CLSID_SensorManager, NULL, CLSCTX_INPROC_SERVER, IID_PPV_ARGS(&pSensorManager));

    if (SUCCEEDED(hr))
        hr = pSensorManager->GetSensorsByCategory(SENSOR_CATEGORY_ALL, &pSensorCollection);

    if (SUCCEEDED(hr))
    {
        ShowCursor(FALSE);

        ULONG count = 0;
        hr = pSensorCollection->GetCount(&count);

        if (count == 0)
        {
            cout << "No sensors found." << endl;
            pSensorCollection->Release();
            pSensorManager->Release();
            CoUninitialize();
			return 0;
        }

        int ulIndex = -1;

        ISensor* pSensor = NULL;
        while (ulIndex < 0 || ulIndex >= count) {
            cout << "Found " << count << " sensors" << endl;

            for (int i = 0; i < count; i++) {
                if (SUCCEEDED(pSensorCollection->GetAt(i, &pSensor))) {
                    BSTR pSensorName;
                    hr = pSensor->GetFriendlyName(&pSensorName);

                    cout << "  " << i << ": " << ConvertBSTRToString(pSensorName) << endl;
                    SysFreeString(pSensorName);
                }
            }

            cout << "Choose one: ";
            cin >> ulIndex;
        }

        HANDLE hConsoleOutput = GetStdHandle(STD_OUTPUT_HANDLE);
        CONSOLE_SCREEN_BUFFER_INFO lpConsoleScreenBufferInfo;
        COORD dwWriteCoord = { 0, 0 };
		DWORD lpNumberOfItemsWritten = 0;

        CONSOLE_CURSOR_INFO curInfo;
        GetConsoleCursorInfo(hConsoleOutput, &curInfo);
        curInfo.bVisible = FALSE;
        SetConsoleCursorInfo(hConsoleOutput, &curInfo);

        if (!GetConsoleScreenBufferInfo(hConsoleOutput, &lpConsoleScreenBufferInfo)) return -1;

        DWORD nLength = lpConsoleScreenBufferInfo.dwSize.X * lpConsoleScreenBufferInfo.dwSize.Y;
        FillConsoleOutputCharacterW(hConsoleOutput, L' ', nLength, dwWriteCoord, &lpNumberOfItemsWritten);
        FillConsoleOutputAttribute(hConsoleOutput, lpConsoleScreenBufferInfo.wAttributes, nLength, dwWriteCoord, &lpNumberOfItemsWritten);

        if (SUCCEEDED(pSensorCollection->GetAt(ulIndex, &pSensor))) {
            BSTR pFriendlyName;
            SENSOR_ID pSensorID;
            SENSOR_TYPE_ID pSensorType;
            SENSOR_CATEGORY_ID pSensorCategory;

            hr = pSensor->GetID(&pSensorID);
            hr = pSensor->GetType(&pSensorType);
            hr = pSensor->GetCategory(&pSensorCategory);
            hr = pSensor->GetFriendlyName(&pFriendlyName);

            LPOLESTR lpSensorIDStr = nullptr;
            LPOLESTR lpSensorTypeStr = nullptr;
            LPOLESTR lpSensorCategoryStr = nullptr;
            hr = StringFromCLSID(pSensorID, &lpSensorIDStr);
            hr = StringFromCLSID(pSensorType, &lpSensorTypeStr);
            hr = StringFromCLSID(pSensorCategory, &lpSensorCategoryStr);

            IPortableDeviceKeyCollection* pDataFields = NULL;
            if (SUCCEEDED(pSensor->GetSupportedDataFields(&pDataFields))) {

                DWORD numFields = 0;
                hr = pDataFields->GetCount(&numFields);

                while (true) {

                    SetConsoleCursorPosition(hConsoleOutput, dwWriteCoord);

                    wcout << "Sensor " << ulIndex << ": " << pFriendlyName << endl << " ID: " << lpSensorIDStr << endl << " Type: " << lpSensorTypeStr << endl << " Category: " << lpSensorCategoryStr << endl;

                    ISensorDataReport* pDataReport = NULL;
                    if (SUCCEEDED(pSensor->GetData(&pDataReport))) {

                        cout << hex << uppercase << setfill('0');

                        for (DWORD j = 0; j < numFields; j++) {
                            PROPERTYKEY pKey;
                            if (SUCCEEDED(pDataFields->GetAt(j, &pKey))) {

                                PROPVARIANT pValue;
                                PropVariantInit(&pValue);

                                if (SUCCEEDED(pDataReport->GetSensorValue(pKey, &pValue))) {

                                    const BYTE* pBytes = reinterpret_cast<const BYTE*>(&pValue) + sizeof(VARTYPE);
                                    constexpr size_t size = sizeof(PROPVARIANT) - sizeof(VARTYPE);

                                    cout << "    ";

                                    for (size_t i = 0; i < size; ++i)
                                        cout << setw(2) << static_cast<int>(pBytes[i]) << " ";

                                    cout << endl;
                                }
                                hr = PropVariantClear(&pValue);
                            }
                        }
                    }
                    pDataReport->Release();
                }
            }
            pDataFields->Release();
        }
		pSensor->Release();
    }
    pSensorCollection->Release();
    pSensorManager->Release();
}

// Run program: Ctrl + F5 or Debug > Start Without Debugging menu
// Debug program: F5 or Debug > Start Debugging menu

// Tips for Getting Started: 
//   1. Use the Solution Explorer window to add/manage files
//   2. Use the Team Explorer window to connect to source control
//   3. Use the Output window to see build output and other messages
//   4. Use the Error List window to view errors
//   5. Go to Project > Add New Item to create new code files, or Project > Add Existing Item to add existing code files to the project
//   6. In the future, to open this project again, go to File > Open > Project and select the .sln file
