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
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_9.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_10.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_11.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_12.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_13.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_14.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_15.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_16.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_17.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_18.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_19.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_20.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_21.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_22.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_23.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_24.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_25.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_26.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_27.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_28.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_29.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_30.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_31.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_32.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_33.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_34.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_35.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_36.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_37.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_38.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_39.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_40.d",
		"C:\\MHDAC_MIDAC_Package\\ExampleData\\rhodamine_test_image_figure4_41.d"
	};

	for (int path=0; path<sizeof(filePaths)/sizeof(*filePaths); path++) {
		CComPtr<IMsdrDataReader> pMSDataReader;
		hr = CoCreateInstance( CLSID_MassSpecDataReader, NULL, CLSCTX_INPROC_SERVER,	
				IID_IMsdrDataReader, (void**)&pMSDataReader);
		assert (hr == S_OK);

		VARIANT_BOOL pRetVal = VARIANT_TRUE;
		hr = pMSDataReader->OpenDataFile(filePaths[path], &pRetVal);
		assert (hr == S_OK);

		CComPtr<BDA::IBDAChromData> pChromData;	
		hr = pMSDataReader ->GetTIC(&pChromData);
		assert (hr == S_OK);

		long dataPoints = 0;
		hr = pChromData->get_TotalDataPoints(&dataPoints);
		assert (hr == S_OK);
		//cout << "Number of data points: " << dataPoints << endl << endl;

		if (path < 10) {
			cout << 0;
		}
		cout << path << ": ";

		for(int scan=0; scan < dataPoints && scan < 41; scan++) {

			CComPtr<IBDASpecFilter> specFilter;

			hr = CoCreateInstance( CLSID_BDAChromFilter, NULL, 
											CLSCTX_INPROC_SERVER ,	
											IID_IBDAChromFilter, 
											(void**)&specFilter);
			if ( S_OK != hr )
			{
				//cout << "You broke it." << endl;
			}

			specFilter->put_SpectrumType(SpecType::SpecType_MassSpectrum);
			specFilter->put_DesiredMSStorageType(DesiredMSStorageType_Profile);

			CComPtr<IMSScanRecord> outputData = NULL;
			hr = pMSDataReader ->GetScanRecord(scan, &outputData);
			assert (hr == S_OK);
			outputData;

			double retentionTime = 0;
			hr = outputData ->get_retentionTime(&retentionTime);
			assert (hr == S_OK);
			MSScanType scanType;
			hr = outputData ->get_MSScanType(&scanType);
			assert (hr == S_OK);
			IonPolarity ionPolarity;
			hr = outputData ->get_ionPolarity(&ionPolarity);
			assert (hr == S_OK);
			IonizationMode ionizationMode;
			hr = outputData ->get_IonizationMode(&ionizationMode);
			assert (hr == S_OK);
			//cout << "Retention time = " << retentionTime << endl;
			//cout << "Scan type = " << scanType << endl;
			//cout << "Ion Polarity = " << ionPolarity << endl;
			//cout << "Ionization mode = " << ionizationMode << endl;
			CComPtr<IBDASpecData> spectrum;
			hr = pMSDataReader ->GetSpectrum(retentionTime, scanType, ionPolarity, ionizationMode, NULL, &spectrum);
			if (hr == S_OK) {
				//cout << endl << "Spectrum loaded successfully." << endl << endl;
				cout << 'Y';

			} else {
				//cout << endl << "Spectrum failed to load. Check source file and try again." << endl << endl;
				cout << 'N';
				continue;
			}

			//long dataPoints = 0;
			hr = spectrum->get_TotalDataPoints(&dataPoints);
			assert (hr == S_OK);
			//cout << "Number of data points: " << dataPoints << endl << endl;

			double* xArray = NULL;
			SAFEARRAY *safeXArray = NULL;
			hr = spectrum->get_XArray(&safeXArray);
			assert (hr == S_OK);
			SafeArrayAccessData(safeXArray, reinterpret_cast<void**>(&xArray));
			SafeArrayUnaccessData(safeXArray);
			//cout << "X Array: ";
			for(int i=0; i<=dataPoints && i<=100; i++) {
				//cout << xArray[i] << ", ";
			}
			//cout << endl << endl;

			double* yArray = NULL;
			SAFEARRAY *safeYArray = NULL;
			hr = spectrum->get_YArray(&safeYArray);
			assert (hr == S_OK);
			SafeArrayAccessData(safeYArray, reinterpret_cast<void**>(&yArray));
			SafeArrayUnaccessData(safeYArray);
			//cout << "Y Array: ";
			for(int i=0; i<=dataPoints && i<=100; i++) {
				//cout << yArray[i] << ", ";
			}
			//cout << endl << endl;

			IRange* timeRange;
			SAFEARRAY *safeTimeArray = NULL;
			hr = pChromData->get_AcquiredTimeRange(&safeTimeArray);
			assert (hr == S_OK);
			SafeArrayAccessData(safeTimeArray, reinterpret_cast<void**>(&timeRange));
			SafeArrayUnaccessData(safeTimeArray);

			//cout << path << ":" << scan << endl;

			//cout << "--------------------------------------------------" << endl << endl;
		}

		cout << endl;
	}

	system("Pause");
	return 0;
}