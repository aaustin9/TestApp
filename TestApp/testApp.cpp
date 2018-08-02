#include <assert.h>
#include <atlbase.h>
#include <filesystem>
#include <iostream>
#include <shobjidl.h>
#include <string>
#include <vector>
#include <windows.h>
#include "mat.h"

#import "BaseCommon.tlb" raw_interfaces_only, no_namespace, named_guids
#import "BaseDataAccess.tlb" raw_interfaces_only, rename_namespace("BDA"), named_guids
#import "MassSpecDataReader.tlb" raw_interfaces_only, no_namespace, named_guids

namespace fs = std::experimental::filesystem;

std::vector<std::wstring> getFoldersFromUserSpecifiedDirectory();
CComBSTR* sortFoldersByNumberInPath(std::vector<std::wstring> list, int size);
bool writeArrayToMATFile(mwSize dimensions[], int numberOfDimensions, double* vals, char variableName[], MATFile &pmat);
double* generateAxisArray(int length, int pixelSize);

int main()
{
	CoInitialize(NULL);
	HRESULT hr = S_OK;

	mwSize current = 0, pixelSizeX = 0, pixelSizeY = 0, xSize = 0, ySize = 0, zSize = 0;
	LONG lBound, uBound, count;
	double *matrix, *zVals;
	bool matrixDefined = false;
	std::vector<float> v;

	std::cout << "Enter integer lengths for pixel width and pixel height." << std::endl;
	std::cout << "Then, select the directory to read spectral data from." << std::endl;
	std::cout << "Enter pixel width: ";
	std::cin >> pixelSizeX;
	std::cout << "Enter pixel height: ";
	std::cin >> pixelSizeY;

	std::vector<std::wstring> spectrumFiles = getFoldersFromUserSpecifiedDirectory();
	ySize = spectrumFiles.size();
	CComBSTR *filePaths = sortFoldersByNumberInPath(spectrumFiles, ySize);

	for (int path = 0; path < ySize; path++) {
		std::cout << "Reading path " << path + 1 << " of " << ySize << "\r";
		CComPtr<IMsdrDataReader> pMSDataReader;
		hr = CoCreateInstance(CLSID_MassSpecDataReader, NULL, CLSCTX_INPROC_SERVER,
				IID_IMsdrDataReader, (void**)&pMSDataReader);
		assert (hr == S_OK);

		VARIANT_BOOL pRetVal = VARIANT_TRUE;
		hr = pMSDataReader->OpenDataFile(filePaths[path], &pRetVal);
		assert (hr == S_OK);

		CComPtr<BDA::IBDAChromData> pChromData;
		hr = pMSDataReader->GetTIC(&pChromData);
		assert (hr == S_OK);

		long dataPoints = 0;
		hr = pChromData->get_TotalDataPoints(&dataPoints);
		if (xSize == 0)
			xSize = dataPoints;
		assert (hr == S_OK);

		for (int scan = 0; scan < xSize; scan++) {

			v.clear();

			if (scan < dataPoints) {
				CComPtr<BDA::IBDASpecFilter> specFilter;

				hr = CoCreateInstance(BDA::CLSID_BDAChromFilter, NULL,
					CLSCTX_INPROC_SERVER,
					BDA::IID_IBDAChromFilter,
					(void**)&specFilter);
				assert (hr == S_OK);

				CComPtr<BDA::IBDASpecData> spectrum;
				hr = pMSDataReader->GetSpectrum_6(scan, NULL, NULL, &spectrum);

				if (hr == S_OK) {
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
						zSize = v.size();
						ySize -= path;
						matrix = new double[ySize * xSize * zSize];
						matrixDefined = true;

						double* xArray = NULL;
						SAFEARRAY *safeXArray = NULL;
						hr = spectrum->get_XArray(&safeXArray);
						assert (hr == S_OK);
						SafeArrayGetLBound(safeXArray, 1, &lBound);
						SafeArrayGetUBound(safeXArray, 1, &uBound);
						SafeArrayAccessData(safeXArray, reinterpret_cast<void**>(&xArray));
						zVals = new double[zSize];
						std::copy(xArray, xArray + zSize, zVals);
					}

				}
			}

			if (matrixDefined) {
				v.resize(zSize, 0);
				copy(v.begin(), v.end(), matrix + current);
				current += zSize;
			}
		}
	}

	std::cout << "Finished reading .d files. Generating MATLAB object." << std::endl;
	std::cout << "This may take several minutes to complete." << std::endl;

	MATFile *pmat;
	mxArray *img, *imgX, *imgY, *imgZ;

	const char *file = "mattest.mat";
	int status;
	double *xVals, *yVals, *start_of_pr;
	double currentVal;
	size_t bytes_to_copy;

	pmat = matOpen(file, "w");
	assert (pmat != NULL);
	assert (writeArrayToMATFile(new mwSize[3]{ zSize, xSize, ySize }, 3, matrix, "img", *pmat));
	std::cout << "Finished processing intensities. Finalizing output." << std::endl;

	assert (writeArrayToMATFile(new mwSize[2]{ 1, xSize }, 2, generateAxisArray(xSize, pixelSizeX), "imgX", *pmat));
	assert (writeArrayToMATFile(new mwSize[2]{ 1, ySize }, 2, generateAxisArray(ySize, pixelSizeY), "imgY", *pmat));
	assert (writeArrayToMATFile(new mwSize[2]{ zSize, 1 }, 2, zVals, "imgZ", *pmat));
	assert (matClose(pmat) == 0);

	std::cout << "Done." << std::endl;
	system("Pause");
	return 0;
}

std::vector<std::wstring> getFoldersFromUserSpecifiedDirectory() {
	std::vector<std::wstring> folders;
	IFileDialog *pfd = NULL;
	DWORD dwOptions;
	IShellItem *psiResult;
	PWSTR pszFilePath = NULL;
	HRESULT hr = CoCreateInstance(CLSID_FileOpenDialog,
		NULL,
		CLSCTX_INPROC_SERVER,
		IID_PPV_ARGS(&pfd));
	assert(SUCCEEDED(hr));

	if (SUCCEEDED(pfd->GetOptions(&dwOptions)))
		pfd->SetOptions(dwOptions | FOS_PICKFOLDERS);
	hr = pfd->Show(NULL);
	assert(SUCCEEDED(hr));

	hr = pfd->GetResult(&psiResult);
	assert(SUCCEEDED(hr));
	hr = psiResult->GetDisplayName(SIGDN_FILESYSPATH, &pszFilePath);
	assert(SUCCEEDED(hr));

	std::wstring ws(pszFilePath);
	std::wstring directoryPath(ws.begin(), ws.end()), directoryContent;

	for (const auto & p : fs::directory_iterator(directoryPath)) {
		directoryContent = p.path().wstring();
		if (directoryContent.substr(directoryContent.length() - 2) == L".d") {
			folders.push_back(directoryContent);
		}
	}
	return folders;
}

CComBSTR* sortFoldersByNumberInPath(std::vector<std::wstring> list, int size) {
	CComBSTR* folders = new CComBSTR[size];
	int position, index;
	for (std::wstring p : list) {
		position = p.length() - 3;
		while (p[position] >= '0' && p[position] <= '9') {
			position--;
		}
		index = stoi(p.substr(position + 1, p.length() - 3 - position));
		folders[index - 1] = SysAllocStringLen(p.data(), p.size());
	}

	return folders;
}

bool writeArrayToMATFile(mwSize dimensions[], int numberOfDimensions, double* vals, char variableName[], MATFile &pmat) {
	mxArray* img = mxCreateNumericArray(numberOfDimensions, dimensions, mxDOUBLE_CLASS, mxREAL);
	assert (img != NULL);
	double* start_of_pr = (double *)mxGetPr(img);
	int bytes_to_copy = mxGetElementSize(img);
	for (int i = 0; i < numberOfDimensions; i++) {
		bytes_to_copy *= dimensions[i];
	}
	memcpy(start_of_pr, vals, bytes_to_copy);
	int status = matPutVariableAsGlobal(&pmat, variableName, img);
	delete dimensions;
	delete vals;
	mxDestroyArray(img);
	return status == 0;
}

double* generateAxisArray(int length, int pixelSize) {
	double* vals = new double[length];
	double currentVal = pixelSize / 2;
	for (int i = 0; i<length; i++) {
		vals[i] = currentVal;
		currentVal += pixelSize;
	}
	return vals;
}