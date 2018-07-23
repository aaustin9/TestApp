#include <assert.h>
#include <atlbase.h>
#include <fstream>
#include <iostream>
#include <vector>
#include <windows.h>
#include <shobjidl.h> 
#include "mat.h"

#import "BaseCommon.tlb" raw_interfaces_only, no_namespace, named_guids
#import "BaseDataAccess.tlb" raw_interfaces_only, rename_namespace("BDA"), named_guids
#import "MassSpecDataReader.tlb" raw_interfaces_only, no_namespace, named_guids

using namespace std;
using namespace BDA;
int main()
{
	CoInitialize(NULL);
	
	int pixelSizeX = 80, pixelSizeY = 80;
	int xSize = 41, ySize = 41, zSize = 0;
	int current = 0, maxLength = 0;

	LONG lBound, uBound, count;

	double * matrix, * zVals;
	bool matrixDefined = false;
	vector<float> v;

	HRESULT hr = S_OK;
	//DIR * dir;
	//IFileDialog *pfd = NULL;
 //   HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, 
 //                     NULL, 
 //                     CLSCTX_INPROC_SERVER, 
 //                     IID_PPV_ARGS(&pfd));
	//assert (SUCCEEDED(hr));
	//DWORD dwOptions;
 //   if (SUCCEEDED(pfd->GetOptions(&dwOptions))) {
 //       pfd->SetOptions(dwOptions | FOS_PICKFOLDERS);
 //   }
	//hr = pfd->Show(NULL);
	//assert (SUCCEEDED(hr));
	//IShellItem *psiResult;
 //   hr = pfd->GetResult(&psiResult);
	//assert (SUCCEEDED(hr));
	//PWSTR pszFilePath = NULL;
	//hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
	//assert (SUCCEEDED(hr));
	//assert ((dir = opendir ("c:\\src\\")) != NULL);

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
			assert (hr == S_OK);

			CComPtr<IBDASpecData> spectrum;
			hr = pMSDataReader ->GetSpectrum_6(scan, NULL, NULL, &spectrum);

			v.clear();

			if (hr == S_OK) {
				cout << 'Y';
				hr = spectrum->get_TotalDataPoints(&dataPoints);
				assert (hr == S_OK);

				float* yArray = NULL;
				SAFEARRAY *safeYArray = NULL;
				hr = spectrum->get_YArray(&safeYArray);
				assert (hr == S_OK);
				SafeArrayGetLBound(safeYArray, 1, &lBound);
				SafeArrayGetUBound(safeYArray, 1, &uBound);
				SafeArrayAccessData(safeYArray, reinterpret_cast<void**>(&yArray));

				v.assign(yArray, yArray + uBound - lBound + 1);
				SafeArrayUnaccessData(safeYArray);

				if (!matrixDefined) {
					assert (scan == 0);
					maxLength = v.size();
					matrix = new double[(41-path) * 41 * maxLength];
					ySize = 41-path;
					zSize = maxLength;
					matrixDefined = true;

					double* xArray = NULL;
					SAFEARRAY *safeXArray = NULL;
					hr = spectrum->get_XArray(&safeXArray);
					assert (hr == S_OK);
					SafeArrayGetLBound(safeXArray, 1, &lBound);
					SafeArrayGetUBound(safeXArray, 1, &uBound);
					SafeArrayAccessData(safeXArray, reinterpret_cast<void**>(&xArray));
					zVals = new double[zSize];
					copy(xArray, xArray + maxLength, zVals);
				}

			} else {
				cout << 'N';
			}

			v.resize(maxLength, 0);
			copy(v.begin(), v.end(), matrix + current);
			current += maxLength;
		}

		cout << endl;
	}



	MATFile *pmat;
    mxArray *img, *imgX, *imgY, *imgZ;

	const char *file = "mattest.mat";
	char str[256];
	int status;
	double * xVals, * yVals, currentVal;
	double * start_of_pr;
	size_t bytes_to_copy;

	pmat = matOpen(file, "w");
	assert (pmat != NULL);

	mwSize dims[] = {zSize, xSize, ySize};
	img = mxCreateNumericArray(3, dims, mxDOUBLE_CLASS, mxREAL);
	assert (img != NULL);
	start_of_pr = (double *)mxGetPr(img);
	bytes_to_copy = xSize * ySize * zSize * mxGetElementSize(img);
	memcpy(start_of_pr, matrix, bytes_to_copy);
	status = matPutVariableAsGlobal(pmat, "img", img);
	assert (status == 0);
	delete matrix;

	mwSize dimsX[] = {1, xSize};
	imgX = mxCreateNumericArray(2,dimsX, mxDOUBLE_CLASS, mxREAL);
	assert (imgX != NULL);
	xVals = new double[xSize];
	currentVal = pixelSizeX/2;
	for(int i=0; i<xSize; i++) {
		xVals[i] = currentVal;
		currentVal += pixelSizeX;
		cout << i << " " << xVals[i] << endl;
	}
	start_of_pr = (double *)mxGetPr(imgX);
	bytes_to_copy = xSize * mxGetElementSize(imgX);
	memcpy(start_of_pr, xVals, bytes_to_copy);
	status = matPutVariableAsGlobal(pmat, "imgX", imgX);
	assert (status == 0);
	delete xVals;

	mwSize dimsY[] = {1, ySize};
	imgY = mxCreateNumericArray(2,dimsY, mxDOUBLE_CLASS, mxREAL);
	assert (imgY != NULL);
	yVals = new double[ySize];
	currentVal = pixelSizeY/2;
	for(int i=0; i<ySize; i++) {
		yVals[i] = currentVal;
		currentVal += pixelSizeY;
	}
	start_of_pr = (double *)mxGetPr(imgY);
	bytes_to_copy = ySize * mxGetElementSize(imgY);
	memcpy(start_of_pr, yVals, bytes_to_copy);
	status = matPutVariableAsGlobal(pmat, "imgY", imgY);
	assert (status == 0);
	delete yVals;

	mwSize dimsZ[] = {zSize, 1};
	imgZ = mxCreateNumericArray(2, dimsZ, mxDOUBLE_CLASS, mxREAL);
	assert (imgZ != NULL);
	start_of_pr = (double *)mxGetPr(imgZ);
	bytes_to_copy = zSize * mxGetElementSize(imgZ);
	memcpy(start_of_pr, zVals, bytes_to_copy);
	status = matPutVariableAsGlobal(pmat, "imgZ", imgZ);
	assert (status == 0);
	delete zVals;

	mxDestroyArray(img);
	mxDestroyArray(imgX);
	mxDestroyArray(imgY);
	mxDestroyArray(imgZ);

	assert (matClose(pmat) == 0);

	system("Pause");
	return 0;
}