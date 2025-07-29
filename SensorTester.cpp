// SensorTester.cpp : This file contains the 'main' function. Program execution begins and ends there.
//

#include <iostream>
#include <string>

#include <Windows.h>
#include <comutil.h>
#include <sensorsapi.h>
#include <sensors.h>
#pragma comment(lib, "comsuppw.lib")
#pragma comment(lib, "sensorsapi.lib")

using namespace std;
using namespace _com_util;

ISensorDataReport* pReport = NULL;
ISensor* pAccelerometer = NULL;

int main()
{
    std::cout << "Sensors:" << endl;

    ISensorManager* pSensorManager = NULL;
    ISensorCollection* pSensorCollection = NULL;
    ISensor* pSensor = NULL;

    HRESULT hr = S_OK;
    // Initialize the sensor manager
    hr = CoInitializeEx(NULL, COINIT_MULTITHREADED);
    if (SUCCEEDED(hr))
    {
        hr = CoCreateInstance(CLSID_SensorManager, NULL, CLSCTX_INPROC_SERVER,
            IID_PPV_ARGS(&pSensorManager));
    }

    // Get accelerometer sensors
    if (SUCCEEDED(hr))
    {
        hr = pSensorManager->GetSensorsByCategory(SENSOR_CATEGORY_ALL, &pSensorCollection);
        //hr = pSensorManager->GetSensorsByCategory(SENSOR_CATEGORY_MOTION, &pSensorCollection);
        //hr = pSensorManager->GetSensorsByType(SENSOR_TYPE_ACCELEROMETER_3D, &pSensorCollection);
        //hr = pSensorManager->GetSensorsByType(/*SENSOR_CATEENGORY_MOTION*/, &pSensorCollection);
    }

    // Get the first accelerometer
    if (SUCCEEDED(hr))
    {
        ULONG count = 0;
        pSensorCollection->GetCount(&count);


        if (count > 0)
        {
            int number = -1;

            while (number < 0 || number >= count) {
                cout << "Found " << count << "sensors" << endl;
                
                for (int i = 0; i < count; i++) {
                    if (SUCCEEDED(pSensorCollection->GetAt(i, &pSensor))) {
                        BSTR name;
                        pSensor->GetFriendlyName(&name);
                        cout << "  " << i << ": " << ConvertBSTRToString(name) << endl;
                        SysFreeString(name);
                    }
				}

                cout << "Choose one: ";
                cin >> number;
            }

            while (true) {
                HANDLE hStdOut = GetStdHandle(STD_OUTPUT_HANDLE);
                COORD coord = { 0, 0 };
                DWORD count;
                CONSOLE_SCREEN_BUFFER_INFO csbi;

                if (GetConsoleScreenBufferInfo(hStdOut, &csbi)) {
                    DWORD cellCount = csbi.dwSize.X * csbi.dwSize.Y;

                    FillConsoleOutputCharacter(hStdOut, ' ', cellCount, coord, &count);

                    FillConsoleOutputAttribute(hStdOut, csbi.wAttributes, cellCount, coord, &count);

                    SetConsoleCursorPosition(hStdOut, coord);
                }

                if (SUCCEEDED(pSensorCollection->GetAt(number, &pSensor))) {
                    BSTR name;
                    pSensor->GetFriendlyName(&name);

                    cout << "  Sensor " << number << ": " << ConvertBSTRToString(name) << endl;

                    SENSOR_CATEGORY_ID categoryId;
                    SENSOR_TYPE_ID typeId;

                    pSensor->GetCategory(&categoryId);
                    pSensor->GetType(&typeId);

                    IPortableDeviceKeyCollection* pDataFields = NULL;
                    if (SUCCEEDED(pSensor->GetSupportedDataFields(&pDataFields))) {
                        DWORD numFields = 0;
                        pDataFields->GetCount(&numFields);

                        for (DWORD j = 0; j < numFields; j++) {
                            PROPERTYKEY key;
                            if (SUCCEEDED(pDataFields->GetAt(j, &key))) {
                                //cout << to_wstring(key.pid).c_str() << endl;
                            }
                        }
                        pDataFields->Release();
                    }

                    ISensorDataReport* pReport = NULL;
                    if (SUCCEEDED(pSensor->GetData(&pReport))) {
                        OutputDebugString(L"   Current Values:\n");

                        IPortableDeviceKeyCollection* pDataFields = NULL;
                        if (SUCCEEDED(pSensor->GetSupportedDataFields(&pDataFields))) {
                            DWORD numFields = 0;
                            pDataFields->GetCount(&numFields);

                            for (DWORD j = 0; j < numFields; j++) {
                                PROPERTYKEY key;
                                if (SUCCEEDED(pDataFields->GetAt(j, &key))) {
                                    PROPVARIANT var;
                                    PropVariantInit(&var);

                                    if (SUCCEEDED(pReport->GetSensorValue(key, &var))) {
                                        std::wstring valueStr;
                                        switch (var.vt) {
                                        case VT_R4: valueStr = std::to_wstring(var.fltVal); break;
                                        case VT_UI4: valueStr = std::to_wstring(var.ulVal); break;
                                        case VT_BOOL: valueStr = var.boolVal ? L"true" : L"false"; break;
                                        default: valueStr = L"[unsupported type]"; break;
                                        }
                                        cout << "    Value: " << valueStr.c_str() << endl;
                                    }
                                    PropVariantClear(&var);
                                }
                            }
                            pDataFields->Release();
                        }
                        pReport->Release();
                    }
                }
            }
        }
    }
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
