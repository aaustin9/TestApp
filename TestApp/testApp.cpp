#include <assert.h>
#include <atlbase.h>
#include <fstream>
#include <iostream>
#include <vector>
#include <windows.h>
#include <shobjidl.h> 
#include <string>
#include <filesystem>
#include "mat.h"

#import "BaseCommon.tlb" raw_interfaces_only, no_namespace, named_guids
#import "BaseDataAccess.tlb" raw_interfaces_only, rename_namespace("BDA"), named_guids
#import "MassSpecDataReader.tlb" raw_interfaces_only, no_namespace, named_guids

using namespace std;
using namespace BDA;
namespace fs = std::experimental::filesystem;

int main()
{
	CoInitialize(NULL);

	
	int pixelSizeX = 0, pixelSizeY = 0;
	int xSize = 0, ySize = 0, zSize = 0;
	int current = 0, maxLength = 0;

	LONG lBound, uBound, count;

	double * matrix, * zVals;
	bool matrixDefined = false;
	vector<float> v;
	vector<wstring> spectrumFiles;
	CComBSTR * filePaths;
	
	cout << "Enter integer lengths for pixel width and pixel height." << endl;
	cout << "Then, select the directory to read spectral data from." << endl;
	cout << "Enter pixel width: ";
	cin >> pixelSizeX;
	cout << "Enter pixel height: ";
	cin >> pixelSizeY;
	IFileDialog *pfd = NULL;
    HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog, 
                      NULL, 
                      CLSCTX_INPROC_SERVER, 
                      IID_PPV_ARGS(&pfd));
	assert (SUCCEEDED(hr));
	DWORD dwOptions;
    if (SUCCEEDED(pfd->GetOptions(&dwOptions)))
        pfd->SetOptions(dwOptions | FOS_PICKFOLDERS);
	hr = pfd->Show(NULL);
	assert (SUCCEEDED(hr));
	IShellItem *psiResult;
    hr = pfd->GetResult(&psiResult);
	assert (SUCCEEDED(hr));
	PWSTR pszFilePath = NULL;
	hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
	assert (SUCCEEDED(hr));
	wstring ws(pszFilePath);
	wstring directoryPath(ws.begin(), ws.end()), directoryContent;
	for (const auto & p : fs::directory_iterator(directoryPath)) {
		directoryContent = p.path().wstring();
		if (directoryContent.substr(directoryContent.length() - 2) == L".d")
			spectrumFiles.push_back(directoryContent);
	}
	int numberOfFiles = spectrumFiles.size();
	ySize = numberOfFiles;
	filePaths = new CComBSTR[numberOfFiles];
	int position, index;
	for (wstring p : spectrumFiles) {
		position = p.length() - 3;
		while (p[position] >= '0' && p[position] <= '9')
			position--;
		index = stoi(p.substr(position + 1, p.length() - 3 - position));
		filePaths[index-1] = SysAllocStringLen(p.data(), p.size());
	}

	for (int path=0; path<numberOfFiles; path++) {
		cout << "Reading path " << path + 1 << " of " << numberOfFiles << "\r";
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

		long spectrumPoints = 0;
		hr = pChromData->get_TotalDataPoints(&spectrumPoints);
		assert (hr == S_OK);

		if (xSize == 0)
			xSize = spectrumPoints;

		for(int scan=0; scan < spectrumPoints && scan < xSize; scan++) {

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
				long dataPoints;
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
					matrix = new double[(ySize-path) * xSize * maxLength];
					ySize = ySize - path;
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

			}
			else {
				cout << "Failed at " << path << " " << scan << endl;
			}

			if (maxLength > 0) {
				v.resize(maxLength, 0);
				copy(v.begin(), v.end(), matrix + current);
				current += maxLength;
			}
		}
	}

	cout << "Finished reading .d files. Generating MATLAB object." << endl;

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

	cout << "Finished processing intensities. Finalizing output." << endl;

	mwSize dimsX[] = {1, xSize};
	imgX = mxCreateNumericArray(2,dimsX, mxDOUBLE_CLASS, mxREAL);
	assert (imgX != NULL);
	xVals = new double[xSize];
	currentVal = pixelSizeX/2;
	for(int i=0; i<xSize; i++) {
		xVals[i] = currentVal;
		currentVal += pixelSizeX;
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

	cout << "Done." << endl;

	system("Pause");
	return 0;
}