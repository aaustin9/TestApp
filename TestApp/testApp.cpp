#include <assert.h>
#include <atlbase.h>
#include <iostream>

#import "BaseCommon.tlb" raw_interfaces_only, no_namespace, named_guids
#import "BaseDataAccess.tlb" raw_interfaces_only, rename_namespace("BDA"), named_guids
#import "MassSpecDataReader.tlb" raw_interfaces_only, no_namespace, named_guids

// static FILE* outputFp = stdout;
using namespace std;
using namespace BDA;
int main()
{
    HRESULT hr = S_OK;
	CoInitialize(NULL);

	CComPtr<IMsdrDataReader> pMSDataReader;
    hr = CoCreateInstance( CLSID_MassSpecDataReader, NULL, CLSCTX_INPROC_SERVER,	
			IID_IMsdrDataReader, (void**)&pMSDataReader);
	assert (hr == S_OK);

	VARIANT_BOOL pRetVal = VARIANT_TRUE;
	CComBSTR filePath = "C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_2.d";
	hr = pMSDataReader->OpenDataFile(filePath, &pRetVal);
	assert (hr == S_OK);

	CComPtr<BDA::IBDAChromData> pChromData;	
	hr = pMSDataReader ->GetTIC(&pChromData);
	assert (hr == S_OK);

	long dataPoints = 0;
	hr = pChromData->get_TotalDataPoints(&dataPoints);
	assert (hr == S_OK);
	cout << dataPoints << endl;

	double* xArray = NULL;
	SAFEARRAY *safeXArray = NULL;
	hr = pChromData->get_XArray(&safeXArray);
	assert (hr == S_OK);
	SafeArrayAccessData(safeXArray, reinterpret_cast<void**>(&xArray));
	SafeArrayUnaccessData(safeXArray);
	for(int i=0; i<=dataPoints; i++) {
		cout << xArray[i] << endl;
	}

	cout << endl << "Break" << endl << endl;

	double* yArray = NULL;
	SAFEARRAY *safeYArray = NULL;
	hr = pChromData->get_YArray(&safeYArray);
	assert (hr == S_OK);
	SafeArrayAccessData(safeYArray, reinterpret_cast<void**>(&yArray));
	SafeArrayUnaccessData(safeYArray);
	for(int i=0; i<=dataPoints; i++) {
		cout << yArray[i] << endl;
	}

	long numPoints = 0;
	hr = pChromData->get_TotalDataPoints(&numPoints);
	assert (hr == S_OK);
	cout << endl << "Total # of points: " << numPoints << endl;

	system("Pause");
	return 0;
}