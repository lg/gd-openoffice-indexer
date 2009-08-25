// LarrysOpenOfficeandStarOfficeIndexerObj.cpp : Implementation of CLarrysOpenOfficeandStarOfficeIndexerObj

#include "stdafx.h"
#include "LarrysOpenOfficeandStarOfficeIndexerObj.h"
#include ".\larrysopenofficeandstarofficeindexerobj.h"

#include "unzip101e\unzip.h"
#include "unzip101e\iowin32.h"

// CLarrysOpenOfficeandStarOfficeIndexerObj

STDMETHODIMP CLarrysOpenOfficeandStarOfficeIndexerObj::InterfaceSupportsErrorInfo(REFIID riid)
{
	static const IID* arr[] = 
	{
		&IID_ILarrysOpenOfficeandStarOfficeIndexerObj
	};

	for (int i=0; i < sizeof(arr) / sizeof(arr[0]); i++)
	{
		if (InlineIsEqualGUID(*arr[i],riid))
			return S_OK;
	}
	return S_FALSE;
}

string GetDataFromContentXML(unzFile UnzipFile) {
	string OutXML;

    // Move to content.xml if it exists inside of the archive
	char ContentFile[256] = "content.xml";
	if (unzLocateFile(UnzipFile, ContentFile, 0) != UNZ_OK) {
		return "";
    }

	// Open the content.xml
	int LastError = UNZ_OK;
	LastError = unzOpenCurrentFilePassword(UnzipFile, 0);
	if (LastError != UNZ_OK) {
		return "";
	}

	// Prepare a buffer to read into
	uInt BufferSize = 8192;
    void* Buffer = (void*)malloc(BufferSize);
    if (Buffer == NULL) {
        return "";
    }

	// Extract the file into memory
	do {
		LastError = unzReadCurrentFile(UnzipFile, Buffer, BufferSize);
		if (LastError < 0) {
			break;
		} else if (LastError > 0) {
			OutXML.append((const char*)Buffer, BufferSize); 
		}
	} while (LastError > 0);

	// Free memory buffer
	free(Buffer);
	if (LastError != UNZ_OK) {
		return "";
	}
	
	return OutXML; 
}

string HTMLToText(string HTML) {
	// The following is a _very simple_ HTML to text convertor. It will simply
	// remove all tags and replace them with new line characters.
	string PlainText;

	size_t CurPos = 0;
	size_t CloseTag = 1;
	size_t OpenTag = 1;

	CloseTag = HTML.find(">", CurPos);
	while ((CloseTag != HTML.npos) && (OpenTag != HTML.npos)) {
		OpenTag = HTML.find("<", CloseTag);
        
		if (OpenTag != HTML.npos) {
			if (OpenTag - CloseTag >= 2) {
				PlainText.append(HTML.substr(CloseTag + 1, OpenTag - CloseTag - 1));
				PlainText.append("\n");
			}

			CurPos = OpenTag;
			CloseTag = HTML.find(">", CurPos);
		}
	}

	return PlainText;
}

STDMETHODIMP CLarrysOpenOfficeandStarOfficeIndexerObj::HandleFile(BSTR RawFullPath, IDispatch* IFactory)
{
	string FullPath = CW2A(RawFullPath);
	
	// Open the zip file (since OpenOffice files are all zips. If you don't
	// believe me, rename one to .zip and decompress it)
	unzFile UnzipFile = NULL;
	zlib_filefunc_def FileFunc;
	fill_win32_filefunc(&FileFunc);
	UnzipFile = unzOpen2(FullPath.c_str(), &FileFunc);
	
	// Verify that the file was successfully opened
	if (UnzipFile == NULL) {
		return S_FALSE;
	}

	// Get data from the Content.xml which is inside of the archive
	string ContentData;
	ContentData = GetDataFromContentXML(UnzipFile);

	// Close the file
	unzClose(UnzipFile);

	if (ContentData == "") {
		return S_FALSE;
	}

	// Convert the XML into plain text
	ContentData = HTMLToText(ContentData);
	
	// Get the modification time of the file
	HANDLE FileHandle;
	FileHandle = CreateFile(FullPath.c_str(), GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, NULL, NULL);

	FILETIME *ModifiedTime = (FILETIME *)malloc(sizeof FILETIME);
	BOOL FileTimeResult = GetFileTime(FileHandle, 0, 0, ModifiedTime);
	CloseHandle(FileHandle);
	if (!FileTimeResult) {
        return S_FALSE; 
	}

	// Format modification date properly
	SYSTEMTIME SystemTime;
	FileTimeToSystemTime(ModifiedTime, &SystemTime);

	// Index file
	LarGDSPlugin GDSPlugin(PluginClassID);
	if (!GDSPlugin.SendTextFileEvent(ContentData, FullPath, SystemTime)) { 
		return S_FALSE;
	}

	return S_OK;	
}
