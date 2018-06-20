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

	CComBSTR filePaths[] = {
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_1.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_2.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_3.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_4.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_5.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_6.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_7.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_8.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_9.d"
	};

	for (int path=0; path<sizeof(filePaths)/sizeof(*filePaths); path++) {
		CComPtr<IMsdrDataReader> pMSDataReader;
		hr = CoCreateInstance( CLSID_MassSpecDataReader, NULL, CLSCTX_INPROC_SERVER,	
				IID_IMsdrDataReader, (void**)&pMSDataReader);
		assert (hr == S_OK);

		__int64 totalDataPoints = 0;
		CComPtr<BDA::IBDAMSScanFileInformation> pScanInfo;
		hr = pMSDataReader->get_MSScanFileInformation(&pScanInfo);
		assert (hr == S_OK);
		hr = pScanInfo->get_TotalScansPresent(&totalDataPoints);
		assert (hr == S_OK);

		cout << "Number of data points (scans): " << totalDataPoints << endl << endl;

	/*	CComPtr<BDA::IBDAMSScanFileInformation> fileInfo;
		hr = pMSDataReader ->get_MSScanFileInformation(&fileInfo);
		assert (hr == S_OK);
		fileInfo;*/

		VARIANT_BOOL pRetVal = VARIANT_TRUE;
		hr = pMSDataReader->OpenDataFile(filePaths[path], &pRetVal);
		assert (hr == S_OK);

		CComPtr<BDA::IBDAChromData> pChromData;	
		hr = pMSDataReader ->GetTIC(&pChromData);
		assert (hr == S_OK);

		long dataPoints = 0;
		hr = pChromData->get_TotalDataPoints(&dataPoints);
		assert (hr == S_OK);
		cout << "Number of data points: " << dataPoints << endl << endl;

		double* xArray = NULL;
		SAFEARRAY *safeXArray = NULL;
		hr = pChromData->get_XArray(&safeXArray);
		assert (hr == S_OK);
		SafeArrayAccessData(safeXArray, reinterpret_cast<void**>(&xArray));
		SafeArrayUnaccessData(safeXArray);
		cout << "X Array: ";
		for(int i=0; i<=dataPoints; i++) {
			cout << xArray[i] << ", ";
		}
		cout << endl << endl;

		//CComPtr<BDA::IBDASpecData> pSpecData;
		//hr = pMSDataReader ->GetSpectrum_6(0, NULL, NULL, &pSpecData);
		//pSpecData;
		//assert (hr == S_OK);

		double* yArray = NULL;
		SAFEARRAY *safeYArray = NULL;
		hr = pChromData->get_YArray(&safeYArray);
		assert (hr == S_OK);
		SafeArrayAccessData(safeYArray, reinterpret_cast<void**>(&yArray));
		SafeArrayUnaccessData(safeYArray);
		cout << "Y Array: ";
		for(int i=0; i<=dataPoints; i++) {
			cout << yArray[i] << ", ";
		}
		cout << endl << endl;

		IRange* timeRange;
		SAFEARRAY *safeTimeArray = NULL;
		hr = pChromData->get_AcquiredTimeRange(&safeTimeArray);
		assert (hr == S_OK);
		SafeArrayAccessData(safeTimeArray, reinterpret_cast<void**>(&timeRange));
		SafeArrayUnaccessData(safeTimeArray);
		/*cout << "Time Range: " << endl;
		double start, end;
		for(int i=0; i<=dataPoints; i++) {
			hr = timeRange[i].get_Start(&start);
			assert (hr == S_OK);
			cout << "\tStart: " << start << endl;
			hr = timeRange[i].get_End(&end);
			assert (hr == S_OK);
			cout << "\tEnd: " << end << endl;
		}
		cout << endl << endl;*/

		//CComPtr<IRange> massRange;
		//SAFEARRAY *safeMassArray = NULL;
		//hr = pChromData->get_MeasuredMassRange(&safeMassArray);
		//assert (hr == S_OK);
		//SafeArrayAccessData(safeMassArray, reinterpret_cast<void**>(&massRange));
		//SafeArrayUnaccessData(safeTimeArray);
		//cout << "Mass Range: " << endl;
		//double start, end;
		//VARIANT_BOOL isEmpty;
		//hr = massRange->IsEmpty(&isEmpty);
		//assert (hr == S_OK);
		//cout << "\tIs Empty: ";
		//if (isEmpty == VARIANT_TRUE) {
		//	cout << "true" << endl;
		//} else {
		//	cout << "false" << endl;
		//}
		//hr = massRange->get_Start(&start);
		//assert (hr == S_OK);
		//cout << "\tStart: " << start << endl;
		//hr = massRange->get_End(&end);
		//assert (hr == S_OK);
		//cout << "\tEnd: " << end << endl;
		//cout << endl << endl;

		long numPoints = 0;
		hr = pChromData->get_TotalDataPoints(&numPoints);
		assert (hr == S_OK);
		cout << "Total # of points: " << numPoints << endl << endl;

		cout << "--------------------------------------------------" << endl << endl;
	}

	system("Pause");
	return 0;
}