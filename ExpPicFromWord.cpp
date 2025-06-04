// 优化版 Word 页眉页脚图片提取工具
#include <windows.h>
#include <comdef.h>
#include <oleauto.h>
#include <shlwapi.h>
#include <gdiplus.h>
#include <iostream>
#include <string>
//#include <algorithm>
//#include <oleauto.h>

#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "gdiplus.lib")

using namespace Gdiplus;

class ComInitializer {
public:
	ComInitializer() {
		CoInitialize(NULL);
		GdiplusStartupInput gdiplusStartupInput;
		GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, NULL);
	}
	~ComInitializer() {
		GdiplusShutdown(gdiplusToken);
		CoUninitialize();
	}
private:
	ULONG_PTR gdiplusToken;
};

int GetEncoderClsid(const wchar_t* format, CLSID* pClsid) {
	UINT num = 0, size = 0;
	GetImageEncodersSize(&num, &size);
	if (size == 0) return -1;
	auto* pImageCodecInfo = (ImageCodecInfo*)(malloc(size));
	if (!pImageCodecInfo) return -1;
	GetImageEncoders(num, size, pImageCodecInfo);
	for (UINT j = 0; j < num; ++j) {
		if (wcscmp(pImageCodecInfo[j].MimeType, format) == 0) {
			*pClsid = pImageCodecInfo[j].Clsid;
			free(pImageCodecInfo);
			return j;
		}
	}
	free(pImageCodecInfo);
	return -1;
}

bool SaveEmfToBitmap(HENHMETAFILE hEmf, const std::wstring& filename, float dpiX = 300.0f, float dpiY = 300.0f)
{
	bool success = false;

	Metafile metafile(hEmf);
	Unit units;
	REAL widthInUnits = metafile.GetWidth();
	REAL heightInUnits = metafile.GetHeight();

	// 将逻辑单位转换为像素（目标 DPI）
	float pixelsPerInch = 100.0f; // EMF 默认单位：0.01 英寸
	int pixelWidth = static_cast<int>((widthInUnits / pixelsPerInch) * dpiX);
	int pixelHeight = static_cast<int>((heightInUnits / pixelsPerInch) * dpiY);

	Bitmap bitmap(pixelWidth, pixelHeight, PixelFormat32bppARGB);
	bitmap.SetResolution(dpiX, dpiY);
	Graphics graphics(&bitmap);
	graphics.Clear(Color::White); // 可选：填白背景
	graphics.DrawImage(&metafile, 0, 0, pixelWidth, pixelHeight);

	CLSID clsid;
	if (GetEncoderClsid(L"image/png", &clsid) != -1) {
		if (bitmap.Save(filename.c_str(), &clsid, nullptr) == Ok) {
			success = true;
		}
	}

	return success;
}

int gImageCount = 0;

bool SaveImageFromClipboard(const std::wstring& filename) {
	GdiplusStartupInput gdiplusStartupInput;
	ULONG_PTR gdiplusToken;
	GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);

	if (!OpenClipboard(nullptr)) return false;

	bool success = false;

	// 1. 尝试 CF_BITMAP
	HANDLE hClipboardData = GetClipboardData(CF_BITMAP);
	if (hClipboardData) {
		Bitmap* pBitmap = new Bitmap((HBITMAP)hClipboardData, nullptr);
		if (pBitmap) {
			CLSID clsid;
			if (GetEncoderClsid(L"image/png", &clsid) != -1 &&
				pBitmap->Save(filename.c_str(), &clsid, nullptr) == Ok) {
				success = true;
			}
			delete pBitmap;
		}
	}

	// 2. 尝试 CF_DIB
	if (!success) {
		hClipboardData = GetClipboardData(CF_DIB);
		if (hClipboardData) {
			BITMAPINFOHEADER* pBmpInfoHeader = (BITMAPINFOHEADER*)GlobalLock(hClipboardData);
			if (pBmpInfoHeader) {
				int width = pBmpInfoHeader->biWidth;
				int height = abs(pBmpInfoHeader->biHeight);  // 高度可能是负的
				int rowStride = ((width * pBmpInfoHeader->biBitCount + 31) / 32) * 4;
				BYTE* pixels = (BYTE*)pBmpInfoHeader + pBmpInfoHeader->biSize;
				Bitmap* pBitmap = new Bitmap(width, height, rowStride, PixelFormat32bppARGB, pixels);
				GlobalUnlock(hClipboardData);

				if (pBitmap) {
					CLSID clsid;
					if (GetEncoderClsid(L"image/png", &clsid) != -1 &&
						pBitmap->Save(filename.c_str(), &clsid, nullptr) == Ok) {
						success = true;
					}
					delete pBitmap;
				}
			}
		}
	}

	// 3. 尝试 CF_ENHMETAFILE（增强型图元文件）
	if (!success) {
		HANDLE hMetaFile = GetClipboardData(CF_ENHMETAFILE);
		if (hMetaFile) {
			HENHMETAFILE hEmf = CopyEnhMetaFile((HENHMETAFILE)hMetaFile, NULL);
			if (hEmf) {
				if (SaveEmfToBitmap(hEmf, filename, 300.0f, 300.0f)) {
					success = true;
				}
				DeleteEnhMetaFile(hEmf);
			}
		}
	}

	CloseClipboard();
	GdiplusShutdown(gdiplusToken);
	return success;
}
Rect GetContentBounds(Bitmap* pBitmap) {
	BitmapData data;
	pBitmap->LockBits(nullptr, ImageLockModeRead, PixelFormat32bppARGB, &data);

	int left = data.Width, right = 0, top = data.Height, bottom = 0;
	BYTE* pPixels = (BYTE*)data.Scan0;

	for (int y = 0; y < data.Height; y++) {
		for (int x = 0; x < data.Width; x++) {
			BYTE alpha = pPixels[3]; // Alpha 通道
			if (alpha > 0) {
				if (x < left) left = x;
				if (x > right) right = x;
				if (y < top) top = y;
				if (y > bottom) bottom = y;
			}
			pPixels += 4;
		}
		pPixels += data.Stride - data.Width * 4;
	}

	pBitmap->UnlockBits(&data);
	return Rect(left, top, right - left + 1, bottom - top + 1);
}

//bool SaveAsEMF(IDispatch* pInlineShape, const std::wstring& filename) {
//    if (!pInlineShape) return false;
//
//    // 1. 获取 EnhMetaFileBits 属性（返回 SAFEARRAY）
//    DISPID dispid = 0;
//    LPOLESTR emfName = _wcsdup(L"EnhMetaFileBits");
//    HRESULT hr = pInlineShape->GetIDsOfNames(IID_NULL, &emfName, 1, LOCALE_USER_DEFAULT, &dispid);
//    free(emfName);
//    if (FAILED(hr)) return false;
//
//    // 2. 调用属性获取方法
//    DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
//    VARIANT result;
//    VariantInit(&result);
//    hr = pInlineShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
//        DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
//    if (FAILED(hr) || result.vt != (VT_ARRAY | VT_UI1)) {
//        VariantClear(&result);
//        return false;
//    }
//
//    // 3. 访问 SAFEARRAY 数据
//    SAFEARRAY* psa = result.parray;
//    BYTE* pData = nullptr;
//    if (FAILED(SafeArrayAccessData(psa, (void**)&pData))) {
//        VariantClear(&result);
//        return false;
//    }
//
//    // 4. 获取 EMF 数据大小
//    LONG lBound = 0, uBound = 0;
//    SafeArrayGetLBound(psa, 1, &lBound);
//    SafeArrayGetUBound(psa, 1, &uBound);
//    DWORD emfSize = uBound - lBound + 1;
//
//    // 5. 直接写入 EMF 数据到文件
//    bool success = false;
//    HANDLE hFile = CreateFileW(
//        filename.c_str(),
//        GENERIC_WRITE,
//        0,
//        nullptr,
//        CREATE_ALWAYS,
//        FILE_ATTRIBUTE_NORMAL,
//        nullptr
//    );
//    if (hFile != INVALID_HANDLE_VALUE) {
//        DWORD bytesWritten = 0;
//        if (WriteFile(hFile, pData, emfSize, &bytesWritten, nullptr)) {
//            success = (bytesWritten == emfSize);
//        }
//        CloseHandle(hFile);
//    }
//
//    // 6. 清理资源
//    SafeArrayUnaccessData(psa);
//    VariantClear(&result);
//    return success;
//}

bool ExtractImageViaEMF(IDispatch* pInlineShape, const std::wstring& filename) {
	// 获取EnhMetaFileBits属性ID
	DISPID dispid = 0;
	LPOLESTR emfName = _wcsdup(L"EnhMetaFileBits");
	HRESULT hr = pInlineShape->GetIDsOfNames(IID_NULL, &emfName, 1, LOCALE_USER_DEFAULT, &dispid);
	free(emfName);
	if (FAILED(hr)) return false;

	// 调用属性获取方法
	DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
	VARIANT result;
	VariantInit(&result);
	hr = pInlineShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
		DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
	if (FAILED(hr) || result.vt != (VT_ARRAY | VT_UI1)) {
		VariantClear(&result);
		return false;
	}

	// 访问安全数组数据
	SAFEARRAY* psa = result.parray;
	BYTE* pData = nullptr;
	if (FAILED(SafeArrayAccessData(psa, (void**)&pData))) {
		VariantClear(&result);
		return false;
	}

	// 创建增强型图元文件
	LONG lBound = 0, uBound = 0;
	SafeArrayGetLBound(psa, 1, &lBound);
	SafeArrayGetUBound(psa, 1, &uBound);
	DWORD size = uBound - lBound + 1;
	HENHMETAFILE hEmf = SetEnhMetaFileBits(size, pData);

	bool success = false;
	if (hEmf) {
		// 获取图元文件头信息
		ENHMETAHEADER emh;
		ZeroMemory(&emh, sizeof(emh));
		if (GetEnhMetaFileHeader(hEmf, sizeof(emh), &emh) == 0) {
			DeleteEnhMetaFile(hEmf);
			SafeArrayUnaccessData(psa);
			VariantClear(&result);
			return false;
		}

		// 转换HIMETRIC单位到像素
		HDC hdcRef = GetDC(nullptr);
		const int dpiX = GetDeviceCaps(hdcRef, LOGPIXELSX);
		const int dpiY = GetDeviceCaps(hdcRef, LOGPIXELSY);
		ReleaseDC(nullptr, hdcRef);

		const LONG frameWidth = emh.rclFrame.right - emh.rclFrame.left;
		const LONG frameHeight = emh.rclFrame.bottom - emh.rclFrame.top;
		const int width = (frameWidth * dpiX + 1270) / 2540;  // 四舍五入
		const int height = (frameHeight * dpiY + 1270) / 2540;

		// 创建目标位图
		Bitmap bitmap(width, height, PixelFormat32bppARGB);
		Graphics graphics(&bitmap);
		graphics.Clear(Color::White);  // 设置白色背景

		// 绘制图元文件
		Metafile metafile(hEmf);
		if (graphics.DrawImage(&metafile, 0, 0, width, height) == Ok) {
			// 查找实际内容边界
			int left = width, right = 0, top = height, bottom = 0;
			for (int y = 0; y < height; ++y) {
				for (int x = 0; x < width; ++x) {
					Gdiplus::Color color;
					bitmap.GetPixel(x, y, &color);
					if (!(color.GetR() == 255 && color.GetG() == 255 && color.GetB() == 255)) {
						if (x < left) left = x;
						if (x > right) right = x;
						if (y < top) top = y;
						if (y > bottom) bottom = y;
					}
				}
			}
			// 确保有内容
			if ((right >= left && bottom >= top)
				&&(right >= left + 10 || bottom >= top + 10) ) { //低于10个像素的图片不要
				int cropWidth = right - left + 1;
				int cropHeight = bottom - top + 1;
				Gdiplus::Bitmap croppedBitmap(cropWidth, cropHeight, PixelFormat32bppARGB);
				Gdiplus::Graphics croppedGraphics(&croppedBitmap);
				croppedGraphics.DrawImage(&bitmap, 0, 0, left, top, cropWidth, cropHeight, Gdiplus::UnitPixel);

				// 保存为PNG
				CLSID clsid;
				if (GetEncoderClsid(L"image/png", &clsid) != -1) {
					if (croppedBitmap.Save(filename.c_str(), &clsid, nullptr) == Gdiplus::Ok) {
						success = true;
					}
				}
			}
			else
			{
				gImageCount--;
				success = true;
			}
		}
		DeleteEnhMetaFile(hEmf);
	}

	// 清理资源
	SafeArrayUnaccessData(psa);
	VariantClear(&result);
	return success;
}

bool TryExtractImageFromInlineShape(IDispatch* pInlineShape, const std::wstring& filename) {
	DISPID dispid;

	VARIANT result;
	VariantInit(&result);
	DISPPARAMS dp = { NULL, NULL, 0, 0 };

	// 方法1：尝试 Range.CopyAsPicture
	OLECHAR* propertyName = _wcsdup(L"Range");
	HRESULT hr = pInlineShape->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	free(propertyName);
	if (FAILED(hr)) {
		return false;
	}

	hr = pInlineShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
	if (FAILED(hr) || result.vt != VT_DISPATCH) {
		VariantClear(&result);
		return false;
	}
	IDispatch* pRange = result.pdispVal;

	LPOLESTR emfName = _wcsdup(L"EnhMetaFileBits");
	hr = pRange->GetIDsOfNames(IID_NULL, &emfName, 1, LOCALE_USER_DEFAULT, &dispid);
	free(emfName);

	if (SUCCEEDED(hr)) {
		if (ExtractImageViaEMF(pRange, filename)) {
			//std::cout << "[EMF] 成功提取图像：" << std::endl;
			return true;
		}
	}

	propertyName = _wcsdup(L"CopyAsPicture");
	hr = pRange->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);

	if (SUCCEEDED(hr)) {
		if (OpenClipboard(nullptr)) {
			EmptyClipboard();
			CloseClipboard();
			hr = pRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
			Sleep(10);
			//int try_count = 0;
			//do
			//{
			if (SaveImageFromClipboard(filename)) {
				//std::cout << "[剪贴板] 成功提取图像：" << std::endl;
				return true;
			}
			//if (try_count != 0)
			//{
			//    Sleep(80);
			//}
   /*         try_count++;
		} while (try_count<2);*/
		}
	}
	pRange->Release();
	free(propertyName);

	//std::cout << "[失败] 无法提取图像：" << std::endl;
	return false;
}

// 递归提取 Range 中所有 InlineShapes（包括表格中的图片）
void ExtractInlineShapesFromRange(IDispatch* pRange, const std::wstring& outputDir, const std::wstring& label, int& imageIndex, int depth = 0) {
	DISPID dispid, dispidCount, dispidItem;
	HRESULT hr;
	VARIANT result;
	DISPPARAMS dp = { nullptr, nullptr, 0, 0 };

	// 获取 InlineShapes 集合
	OLECHAR* propInlineShapes = _wcsdup(L"InlineShapes");
	hr = pRange->GetIDsOfNames(IID_NULL, &propInlineShapes, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) return;

	VariantInit(&result);
	hr = pRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
	if (FAILED(hr) || result.vt != VT_DISPATCH) return;
	IDispatch* pInlineShapes = result.pdispVal;

	// 获取 Count
	OLECHAR* countProp = _wcsdup(L"Count");
	hr = pInlineShapes->GetIDsOfNames(IID_NULL, &countProp, 1, LOCALE_USER_DEFAULT, &dispidCount);
	VariantInit(&result);
	hr = pInlineShapes->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
	if (FAILED(hr)) {
		pInlineShapes->Release();
		return;
	}
	long shapeCount = result.lVal;

	// 获取 Item
	OLECHAR* itemProp = _wcsdup(L"Item");
	hr = pInlineShapes->GetIDsOfNames(IID_NULL, &itemProp, 1, LOCALE_USER_DEFAULT, &dispidItem);

	for (long i = 1; i <= shapeCount; ++i) {
		VARIANT idx;
		idx.vt = VT_I4;
		idx.lVal = i;
		DISPPARAMS dpItem = { &idx, NULL, 1, 0 };
		VariantInit(&result);
		hr = pInlineShapes->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpItem, &result, NULL, NULL);
		if (FAILED(hr) || result.vt != VT_DISPATCH) continue;

		IDispatch* pInlineShape = result.pdispVal;

		wchar_t filename[256];
		swprintf_s(filename, L"%s\\header_%s_%d.png", outputDir.c_str(), label.c_str(), imageIndex++);
		bool re = TryExtractImageFromInlineShape(pInlineShape, filename);
		if (re)
		{
			gImageCount++;
			std::cout << gImageCount << "保存页眉图片成功" << std::endl;
		}
		pInlineShape->Release();
	}
	pInlineShapes->Release();
}

// 辅助函数：获取IDispatch对象的属性
bool GetProperty(IDispatch* pDisp, const OLECHAR* name, VARIANT* pResult) {
	DISPID dispid;
	if (FAILED(pDisp->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&name), 1, LOCALE_USER_DEFAULT, &dispid))) {
		return false;
	}
	DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
	return SUCCEEDED(pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, pResult, nullptr, nullptr));
}

long GetSectionStartPage(IDispatch* pSection) {
	VARIANT result;
	VariantInit(&result);

	auto getProp = [&](IDispatch* obj, const OLECHAR* name, VARIANT* pResult) -> bool {
		DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
		DISPID id;
		if (FAILED(obj->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&name), 1, LOCALE_USER_DEFAULT, &id))) return false;
		return SUCCEEDED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, pResult, nullptr, nullptr));
		};

	VARIANT rangeVar;
	VariantInit(&rangeVar);
	if (!getProp(pSection, L"Range", &rangeVar) || rangeVar.vt != VT_DISPATCH) return -1;
	IDispatch* pRange = rangeVar.pdispVal;

	// .Information(3) = wdActiveEndPageNumber
	DISPID dispidInfo;
	OLECHAR* infoProp = _wcsdup(L"Information");
	VARIANT arg;
	arg.vt = VT_I4;
	arg.lVal = 3;

	if (SUCCEEDED(pRange->GetIDsOfNames(IID_NULL, &infoProp, 1, LOCALE_USER_DEFAULT, &dispidInfo))) {
		DISPPARAMS dpInfo = { &arg, nullptr, 1, 0 };
		if (SUCCEEDED(pRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpInfo, &result, nullptr, nullptr))) {
			pRange->Release();
			return result.lVal;
		}
	}
	pRange->Release();
	return -1;
}

bool TryExtractImageFromShape(IDispatch* pWordApp, IDispatch* pShape, const std::wstring& filename) {
	HRESULT hr;
	DISPID dispid;

	// 1. 选中图形
	OLECHAR* selectMethod = _wcsdup(L"Select");
	hr = pShape->GetIDsOfNames(IID_NULL, &selectMethod, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) return false;

	DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
	hr = pShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &noArgs, NULL, NULL, NULL);
	if (FAILED(hr)) return false;

	// 2. 获取 Selection
	IDispatch* pSelection = nullptr;
	VARIANT selResult;
	VariantInit(&selResult);
	DISPID dispidSelection;
	OLECHAR* selProp = _wcsdup(L"Selection");
	hr = pWordApp->GetIDsOfNames(IID_NULL, &selProp, 1, LOCALE_USER_DEFAULT, &dispidSelection);
	if (FAILED(hr)) return false;

	hr = pWordApp->Invoke(dispidSelection, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &selResult, NULL, NULL);
	if (FAILED(hr) || selResult.vt != VT_DISPATCH) return false;

	pSelection = selResult.pdispVal;

	LPOLESTR emfName = _wcsdup(L"EnhMetaFileBits");
	hr = pSelection->GetIDsOfNames(IID_NULL, &emfName, 1, LOCALE_USER_DEFAULT, &dispid);
	free(emfName);

	if (SUCCEEDED(hr)) {
		if (ExtractImageViaEMF(pSelection, filename)) {
			//std::cout << "[EMF] 成功提取图像：" << std::endl;
			return true;
		}
	}

	// 3. Selection.Copy()
	OLECHAR* copyMethod = _wcsdup(L"CopyAsPicture");
	hr = pSelection->GetIDsOfNames(IID_NULL, &copyMethod, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr)) {
		pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &noArgs, NULL, NULL, NULL);
	}
	pSelection->Release();

	// 4. 从剪贴板提取图像
	Sleep(10);  // 等待剪贴板准备
	return SaveImageFromClipboard(filename);
}



bool TryExtractImageSmart(IDispatch* pWordApp, IDispatch* pShape, const std::wstring& filename) {
	// 尝试方法 1：CopyAsPicture + 剪贴板提取
	if (TryExtractImageFromShape(pWordApp, pShape, filename)) {
		//std::cout << "成功提取图片" << std::endl;
		return true;
	}

	// 尝试方法 2：Export 方法（某些图片/图表/艺术字支持）
	DISPID dispidExport;
	OLECHAR* exportMethod = _wcsdup(L"Export");
	HRESULT hr = pShape->GetIDsOfNames(IID_NULL, &exportMethod, 1, LOCALE_USER_DEFAULT, &dispidExport);
	free(exportMethod);
	if (SUCCEEDED(hr)) {
		VARIANT varFilename, varFormat;
		varFilename.vt = VT_BSTR;
		varFilename.bstrVal = SysAllocString(filename.c_str());

		varFormat.vt = VT_BSTR;
		varFormat.bstrVal = SysAllocString(L"PNG"); // 你也可以试 EMF

		VARIANT args[2] = { varFormat, varFilename };
		DISPPARAMS dpExport = { args, nullptr, 2, 0 };

		hr = pShape->Invoke(dispidExport, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpExport, nullptr, nullptr, nullptr);
		VariantClear(&varFilename);
		VariantClear(&varFormat);

		if (SUCCEEDED(hr)) {
			std::cout << "[Smart] 使用 Export 成功导出图片" << std::endl;
			return true;
		}
	}

	// 尝试方法 3：ConvertToInlineShape → TryExtractImageFromInlineShape
	DISPID dispidConvert;
	OLECHAR* convertMethod = _wcsdup(L"ConvertToInlineShape");
	hr = pShape->GetIDsOfNames(IID_NULL, &convertMethod, 1, LOCALE_USER_DEFAULT, &dispidConvert);
	free(convertMethod);

	if (SUCCEEDED(hr)) {
		DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
		VARIANT result;
		VariantInit(&result);
		hr = pShape->Invoke(dispidConvert, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &noArgs, &result, NULL, NULL);
		if (SUCCEEDED(hr) && result.vt == VT_DISPATCH && result.pdispVal) {
			IDispatch* pInlineShape = result.pdispVal;
			bool success = false;
			for (size_t i = 0; i < 5; i++)
			{
				success = TryExtractImageFromInlineShape(pInlineShape, filename);
				if (success)
				{
					break;
				}
			}
			pInlineShape->Release();
			if (success) {
				std::cout << "[Smart] 转换为内联形状后提取成功" << std::endl;
				return true;
			}
		}
	}

	std::cout << "[Smart] 所有提取方式均失败" << std::endl;
	return false;
}



int ExportFloatingShapes(IDispatch* pDocument, IDispatch* pWordApp, const std::wstring& outputDir) {
	HRESULT hr;
	DISPID dispid;
	DISPPARAMS noArgs = { NULL, NULL, 0, 0 };
	VARIANT result;
	IDispatch* pShapes = nullptr;

	// 获取 Shapes
	OLECHAR* propShapes = _wcsdup(L"Shapes");
	hr = pDocument->GetIDsOfNames(IID_NULL, &propShapes, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) return -1;

	hr = pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr) || result.vt != VT_DISPATCH) return -1;
	pShapes = result.pdispVal;

	// 获取 Shapes.Count
	DISPID dispidCount;
	OLECHAR* countProp = _wcsdup(L"Count");
	hr = pShapes->GetIDsOfNames(IID_NULL, &countProp, 1, LOCALE_USER_DEFAULT, &dispidCount);
	if (FAILED(hr)) return -1;

	VariantInit(&result);
	hr = pShapes->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
	if (FAILED(hr)) return -1;
	long shapeCount = result.lVal;

	// 获取 Shapes.Item 方法
	DISPID dispidItem;
	OLECHAR* itemMethod = _wcsdup(L"Item");
	hr = pShapes->GetIDsOfNames(IID_NULL, &itemMethod, 1, LOCALE_USER_DEFAULT, &dispidItem);
	if (FAILED(hr)) return -1;

	static int float_try_count;
	int imagePageIndex = 1, lastPage = 0;// 上次页码  和本页图片
	for (long i = 1; i <= shapeCount; ++i) {
		float_try_count = 0;
	gofloattry:
		VARIANT index;
		index.vt = VT_I4;
		index.lVal = i;
		DISPPARAMS dpItem = { &index, NULL, 1, 0 };
		VariantInit(&result);
		hr = pShapes->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpItem, &result, NULL, NULL);
		if (FAILED(hr) || result.vt != VT_DISPATCH) continue;
		IDispatch* pShape = result.pdispVal;

		// 获取 Anchor.Range.Information(wdActiveEndPageNumber)
		int pageNumber = 0;
		DISPID dispidAnchor, dispidRange, dispidInfo;
		VARIANT anchorVar;
		VariantInit(&anchorVar);
		OLECHAR* anchorProp = _wcsdup(L"Anchor");
		hr = pShape->GetIDsOfNames(IID_NULL, &anchorProp, 1, LOCALE_USER_DEFAULT, &dispidAnchor);
		if (SUCCEEDED(hr)) {
			pShape->Invoke(dispidAnchor, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &noArgs, &anchorVar, NULL, NULL);
			if (anchorVar.vt == VT_DISPATCH) {
				IDispatch* pRange = anchorVar.pdispVal;
				OLECHAR* rangeProp = _wcsdup(L"Information");
				hr = pRange->GetIDsOfNames(IID_NULL, &rangeProp, 1, LOCALE_USER_DEFAULT, &dispidInfo);
				if (SUCCEEDED(hr)) {
					VARIANT arg;
					arg.vt = VT_I4;
					arg.lVal = 3; // wdActiveEndPageNumber
					DISPPARAMS dpInfo = { &arg, NULL, 1, 0 };
					VariantInit(&result);
					hr = pRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dpInfo, &result, NULL, NULL);
					if (SUCCEEDED(hr)) pageNumber = result.lVal;
				}
				pRange->Release();
			}
		}
		if (pageNumber == 0)
		{
			pageNumber = 1;
		}
		if (pageNumber == lastPage)
		{
			imagePageIndex++;
		}
		else
		{
			imagePageIndex = 1;
		}
		// 生成文件名
		wchar_t filename[MAX_PATH];
		swprintf_s(filename, L"%s\\float_image_%d_%d.png", outputDir.c_str(), pageNumber, imagePageIndex);

		if (TryExtractImageSmart(pWordApp, pShape, filename)) {
			gImageCount++;
			std::cout << gImageCount << "已保存浮动图片: " << std::endl;
			lastPage = pageNumber;
		}
		else {
			if (pageNumber == lastPage)
			{
				imagePageIndex--;
			}
			std::cout << gImageCount << "无法保存图片: page = " << pageNumber << std::endl;
			float_try_count++;
			if (float_try_count <= 3)
			{
				goto gofloattry;
			}
		}
		pShape->Release();
	}

	pShapes->Release();
	return 0;
}


int last_section = 1;

void ExtractImagesFromHeaders(IDispatch* pDocument, IDispatch* pWordApp, const std::wstring& outputDir) {
	DISPID dispid;
	VARIANT result;
	DISPPARAMS dp = { nullptr, nullptr, 0, 0 };
	auto getProp = [&](IDispatch* obj, const OLECHAR* name, VARIANT* pResult) -> bool {
		DISPID id;
		if (FAILED(obj->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR*>(&name), 1, LOCALE_USER_DEFAULT, &id))) return false;
		return SUCCEEDED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, pResult, nullptr, nullptr));
		};

	VARIANT sectionsVar;
	VariantInit(&sectionsVar);
	if (!getProp(pDocument, L"Sections", &sectionsVar) || sectionsVar.vt != VT_DISPATCH) return;
	IDispatch* pSections = sectionsVar.pdispVal;

	DISPID itemDispid;
	OLECHAR* itemName = _wcsdup(L"Item");
	pSections->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &itemDispid);

	VARIANT countVar;
	VariantInit(&countVar);
	getProp(pSections, L"Count", &countVar);
	long sectionCount = countVar.lVal;

	int imageIndex = 1;
	for (long i = 1; i <= sectionCount; ++i) {
		VARIANT index;
		index.vt = VT_I4;
		index.lVal = i;
		dp.rgvarg = &index;
		dp.cArgs = 1;
		VariantInit(&result);
		pSections->Invoke(itemDispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, nullptr, nullptr);
		if (!result.pdispVal) continue;
		IDispatch* pSection = result.pdispVal;
		long startPage = GetSectionStartPage(pSection);

		for (const OLECHAR* type : { _wcsdup(L"Headers"), _wcsdup(L"Footers") }) {
			VARIANT hfVar;
			VariantInit(&hfVar);

			dp = { nullptr, nullptr, 0, 0 };
			if (!getProp(pSection, type, &hfVar) || hfVar.vt != VT_DISPATCH) continue;
			IDispatch* pCollection = hfVar.pdispVal;

			getProp(pCollection, L"Count", &countVar);
			long hfCount = countVar.lVal;

			DISPID hfItemDispid;
			pCollection->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &hfItemDispid);

			for (long j = 1; j <= 1; ++j) {
				index.lVal = j;
				dp.rgvarg = &index;
				dp.cArgs = 1;
				VariantInit(&result);
				pCollection->Invoke(hfItemDispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, nullptr, nullptr);
				if (!result.pdispVal) continue;
				IDispatch* pHF = result.pdispVal;

				VARIANT rangeVar;
				VariantInit(&rangeVar);
				dp = { nullptr, nullptr, 0, 0 };
				if (!getProp(pHF, L"Range", &rangeVar) || rangeVar.vt != VT_DISPATCH) continue;
				IDispatch* pRange = rangeVar.pdispVal;

				std::wstring label = std::to_wstring(last_section);
				ExtractInlineShapesFromRange(pRange, outputDir, label, imageIndex);

				static int ccc ;
				if (ccc==0)
				{
					ccc++;
					ExportFloatingShapes(pHF, pWordApp, outputDir);
				}
				//ExportFloatingShapes(pWordApp, pHF, outputDir);
				// 新增提取 遍历浮动 Shape
				// 1. 拿到 Shapes 集合
				//VARIANT shapesVar;
				//VariantInit(&shapesVar);
				//if (SUCCEEDED(getProp(pHF, L"Shapes", &shapesVar)) && shapesVar.vt == VT_DISPATCH) {
				//	IDispatch* pShapes = shapesVar.pdispVal;
				//	// 2. 取 Count
				//	VARIANT countVar; VariantInit(&countVar);
				//	getProp(pShapes, L"Count", &countVar);
				//	long shapeCount = countVar.lVal;

				//	// 3. Item 方法 DISPID
				//	DISPID dispidItem;
				//	OLECHAR* itemName = _wcsdup(L"Item");
				//	pShapes->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT, &dispidItem);

				//	// 4. 遍历
				//	for (long si = 1; si <= shapeCount; ++si) {
				//		VARIANT idx; idx.vt = VT_I4; idx.lVal = si;
				//		DISPPARAMS dpItem = { &idx, NULL, 1, 0 };
				//		VARIANT shpResult; VariantInit(&shpResult);
				//		if (SUCCEEDED(pShapes->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpItem, &shpResult, NULL, NULL))
				//			&& shpResult.vt == VT_DISPATCH) {

				//			IDispatch* pShape = shpResult.pdispVal;
				//			VARIANT typeVar;
				//			getProp(pShape, L"Type", &typeVar);

				//			long shpType = typeVar.lVal;
				//			std::cout << "shp type = " << shpType << std::endl;
				//			// 只处理图片类型

				//			wchar_t shpFile[MAX_PATH];
				//			swprintf_s(shpFile, L"%s\\header_%s_%d.png", outputDir.c_str(), label.c_str(), imageIndex++);

				//			if (TryExtractImageSmart(pWordApp, pShape, shpFile)) {
				//				std::wcout << L"浮动 Shape 保存成功: " << shpFile << std::endl;
				//			}

				//			pShape->Release();
				//		}
				//		VariantClear(&shpResult);
				//	}
				//	pShapes->Release();
				//}

				pRange->Release();
				pHF->Release();
			}
			pCollection->Release();
		}
		last_section = startPage + 1;
		pSection->Release();
	}
	pSections->Release();
}

// Word Application CLSID 和 IID
const CLSID CLSID_WordApplication = { 0x000209FF, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };
const CLSID CLSID_WpsApplication = { 0x000209FF, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };
const IID IID__Application = { 0x00020970, 0x0000, 0x0000, {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46} };

int ExportpInlineShapes(IDispatch* pDocument, DISPID dispid, std::wstring outputDir)
{
	// 获取所有内联形状（包括图片）
	HRESULT hr;
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	IDispatch* pInlineShapes = NULL;
	OLECHAR* propertyName = _wcsdup(L"InlineShapes");
	hr = pDocument->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) {
		pDocument->Release();
		std::cout << "无法获取InlineShapes属性" << std::endl;
		return false;
	}
	//std::cout << "获取InlineShapes属性" << std::endl;
	VARIANT result;
	VariantInit(&result);
	dp.cArgs = 0;
	dp.rgvarg = NULL;
	hr = pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
	if (FAILED(hr)) {
		std::cout << "无法获取InlineShapes集合" << std::endl;
		return false;
	}
	//std::cout << "获取InlineShapes集合" << std::endl;
	pInlineShapes = result.pdispVal;

	// 获取形状数量
	propertyName = _wcsdup(L"Count");
	hr = pInlineShapes->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) {
		pInlineShapes->Release();
		std::cout << "无法获取Count属性" << std::endl;
		return false;
	}
	//std::cout << "获取Count属性" << std::endl;

	long count = 0;
	VariantInit(&result);
	dp.cArgs = 0;
	dp.rgvarg = NULL;
	hr = pInlineShapes->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
	if (SUCCEEDED(hr)) {
		count = result.lVal;
	}

	// 遍历所有形状
	propertyName = _wcsdup(L"Item");
	hr = pInlineShapes->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) {
		pInlineShapes->Release();
		std::cout << "无法获取Item方法" << std::endl;
		return false;
	}

	VariantInit(&result);
	dp = { NULL, NULL, 0, 0 };

	static int try_count;// 尝试次数
	int imagePageIndex = 1, lastPage = 0;// 上次页码  和本页图片
	for (long i = 1; i <= count; i++) {
		gImageCount = i;// 计数
		try_count = 0;
		auto dispid1 = dispid;
		VARIANT varIndex;
		varIndex.vt = VT_I4;
		varIndex.lVal = i;

		dp.cArgs = 1;
		dp.rgvarg = &varIndex;

		VariantInit(&result);
		hr = pInlineShapes->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
		if (FAILED(hr) || result.pdispVal == NULL) {
			_com_error err(hr);
			std::cout << "获取第 " << i << " 个形状失败 (HRESULT: 0x"
				<< std::hex << hr << "): " << err.ErrorMessage() << std::endl;
			VariantClear(&result);
			continue;
		}

		IDispatch* pInlineShape = result.pdispVal;

	gotry:
		IDispatch* pRange = NULL;
		propertyName = _wcsdup(L"Range");
		hr = pInlineShape->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid1);
		if (FAILED(hr)) {
			pInlineShape->Release();
			continue;
		}

		VariantInit(&result);
		dp.cArgs = 0;
		dp.rgvarg = NULL;
		hr = pInlineShape->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
		if (FAILED(hr) || result.pdispVal == NULL) {
			pInlineShape->Release();
			continue;
		}
		pRange = result.pdispVal;

		// 获取页码
		long pageNumber = 0;
		propertyName = _wcsdup(L"Information");
		hr = pRange->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid1);
		if (SUCCEEDED(hr)) {
			VARIANT varWdInformation;
			varWdInformation.vt = VT_I4;
			varWdInformation.lVal = 3; // wdActiveEndPageNumber

			dp.cArgs = 1;
			dp.rgvarg = &varWdInformation;

			VariantInit(&result);
			hr = pRange->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
			if (SUCCEEDED(hr)) {
				pageNumber = result.lVal;
			}
		}
		if (pageNumber == lastPage)
		{
			imagePageIndex++;
		}
		else
		{
			imagePageIndex = 1;
		}
		// 创建输出文件名
		wchar_t filename[MAX_PATH];
		swprintf_s(filename, L"%s\\image_%d_%d.png", outputDir.c_str(), pageNumber, imagePageIndex);

		lastPage = pageNumber;

		// 提取并保存图片
		if (TryExtractImageFromInlineShape(pInlineShape, filename)) {
			std::cout << i << "已保存图片 " << std::endl;
		}
		else {
			std::cout << i << "无法保存图片" << std::endl;
			if (pageNumber == lastPage)
			{
				imagePageIndex--;
			}
			try_count++;
			if (try_count <= 5)
			{
				goto gotry;
			}
			else
			{
				std::cout << i << "--------------------无法保存图片----------------------" << std::endl;
			}
		}
		pRange->Release();
		// 清理资源
		pInlineShape->Release();
		VariantClear(&result);
	}
	return 0;
}

int floatToInline(IDispatch* pWordApp, IDispatch* pDocument)
{
	// 转换非内嵌图片（Shapes）
	IDispatch* pShapes = NULL;
	DISPID dispid;
	HRESULT hr;
	VARIANT result;
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	OLECHAR* propertyName = _wcsdup(L"Shapes");
	hr = pDocument->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr)) {
		VariantInit(&result);
		dp.cArgs = 0;
		hr = pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
		if (SUCCEEDED(hr) && result.pdispVal) {
			pShapes = result.pdispVal;

			// 遍历Shapes
			propertyName = _wcsdup(L"Count");
			hr = pShapes->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
			if (SUCCEEDED(hr)) {
				VariantInit(&result);
				hr = pShapes->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
				long shapesCount = result.lVal;

				propertyName = _wcsdup(L"Item");
				hr = pShapes->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
				if (FAILED(hr)) {
					std::cout << "无法获取Item方法" << std::endl;
					return false;
				}
				auto itemDispid = dispid;
				for (int i = 1; i <= shapesCount; i++) {
					VARIANT varIndex;
					VariantInit(&varIndex);
					varIndex.vt = VT_I4;
					varIndex.lVal = i;
					dp.cArgs = 1;
					dp.rgvarg = &varIndex;

					VariantInit(&result);
					hr = pShapes->Invoke(itemDispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);
					if (FAILED(hr)) {
						_com_error err(hr);
						std::cout << "获取第 " << i << " 个Shape失败: " << err.ErrorMessage() << std::endl;
						continue;
					}
					IDispatch* pShape = result.pdispVal;
					/*        if (ExtractImageViaEMF(pShapes, L"filename")) {
								std::cout << "已保存页眉页脚图片 " << std::endl;
							}
							else {
								std::cout << "无法保存页眉页脚图片" << std::endl;
							}*/

							// 转换为内联形状(图片)
					propertyName = _wcsdup(L"ConvertToInlineShape");
					hr = pShape->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);

					if (SUCCEEDED(hr)) {
						dp = { NULL, NULL, 0, 0 };
						hr = pShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, &result, NULL, NULL);

						if (FAILED(hr)) {
							std::cout << "转换失败" << std::endl;
							continue;
						}
					}
				}
			}
		}
	}
	pShapes->Release();
}

// 递归创建宽字符目录
// 该函数会遍历路径中的每个子目录，如果子目录不存在则创建它
bool CreateDirectoryRecursive(const std::wstring& path) {
	// 如果路径为空，直接返回 false
	if (path.empty()) {
		return false;
	}
	size_t pos = 0;
	std::wstring currentPath;
	// 遍历路径，查找路径分隔符
	while ((pos = path.find_first_of(L"\\/", pos)) != std::wstring::npos) {
		currentPath = path.substr(0, pos++);
		// 如果当前子目录不存在，则尝试创建它
		if (!currentPath.empty() && !PathFileExistsW(currentPath.c_str())) {
			if (!CreateDirectoryW(currentPath.c_str(), NULL)) {
				// 获取错误代码
				//DWORD error = GetLastError();
				//std::wcerr << L"Failed to create directory: " << currentPath << L", Error code: " << error << std::endl;
				return false;
			}
		}
	}
	// 处理最后一个子目录
	if (!path.empty() && !PathFileExistsW(path.c_str())) {
		if (!CreateDirectoryW(path.c_str(), NULL)) {
			// 获取错误代码
			//DWORD error = GetLastError();
			//std::wcerr << L"Failed to create directory: " << path << L", Error code: " << error << std::endl;
			return false;
		}
	}
	return true;
}

#include <chrono>

int wmain(int argc, wchar_t* argv[]) {
#if 0
	if (argc != 3) {
		
		return 1;
	}

	std::wstring docPath = argv[1];
	std::wstring outputDir = argv[2];

#else
	std::wstring docPath = L"F:/66.docx";
	std::wstring outputDir = L"F:/66pic";
#endif

	auto start = std::chrono::high_resolution_clock::now();
	ComInitializer comInit;

	if (!PathFileExistsW(outputDir.c_str())) {
		if (CreateDirectoryRecursive(outputDir)) {
			std::cout << "Directory created successfully." << std::endl;
		}
		else {
			std::cout << "Failed to create directory." << std::endl;
		}
	}
	// 创建Word应用程序对象
	HRESULT hr;
	IDispatch* pWordApp = NULL;
	hr = CoCreateInstance(CLSID_WordApplication, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&pWordApp);
	if (FAILED(hr)) {
		std::cout << "无法创建Word应用程序实例" << std::endl;
		return false;
	}
	std::cout << "创建Word应用程序实例" << std::endl;

	// 设置Word不可见
	DISPID dispid;
	OLECHAR* propertyName = _wcsdup(L"Visible");
	hr = pWordApp->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr)) {
		VARIANT var;
		var.vt = VT_BOOL;
		var.boolVal = VARIANT_FALSE;
		DISPPARAMS dp = { &var, NULL, 1, 0 };
		pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
	}

	// 打开文档
	IDispatch* pDocuments = NULL;
	propertyName = _wcsdup(L"Documents");
	hr = pWordApp->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) {
		pWordApp->Release();
		std::cout << "无法获取Documents属性" << std::endl;
		return false;
	}

	std::cout << "获取Documents属性" << std::endl;
	VARIANT result;
	VariantInit(&result);
	DISPPARAMS dp = { NULL, NULL, 0, 0 };
	hr = pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
	if (FAILED(hr) || result.pdispVal == NULL) {
		pWordApp->Release();
		std::cout << "无法获取Documents集合" << std::endl;
		return false;
	}
	std::cout << "获取Documents集合" << std::endl;
	pDocuments = result.pdispVal;

	// 调用Documents.Open方法
	IDispatch* pDocument = NULL;
	propertyName = _wcsdup(L"Open");
	hr = pDocuments->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr)) {
		pDocuments->Release();
		pWordApp->Release();
		std::cout << "无法获取Open方法" << std::endl;
		return false;
	}

	VARIANT args[1];
	args[0].vt = VT_BSTR;
	args[0].bstrVal = SysAllocString(docPath.c_str());
	DISPPARAMS dpOpen = { args, nullptr, 1, 0 };
	VariantInit(&result);
	hr = pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dpOpen, &result, nullptr, nullptr);
	pDocument = result.pdispVal;

	if (pDocument == nullptr)
	{
		std::cout << "Open Document failed" << std::endl;
		return -1;
	}
	std::cout << "Open Document success" << std::endl;

	ExportpInlineShapes(pDocument, dispid, outputDir);

	std::cout << "====================================" << std::endl;
	ExtractImagesFromHeaders(pDocument,pWordApp, outputDir);
	std::cout << "====================================" << std::endl;
	ExportFloatingShapes(pDocument, pWordApp, outputDir);

	// 关闭文档
	propertyName = _wcsdup(L"Close");
	hr = pDocument->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr)) {
		dp.cArgs = 0;
		dp.rgvarg = NULL;
		pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
	}
	// 退出Word
	propertyName = _wcsdup(L"Quit");
	hr = pWordApp->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr)) {
		dp.cArgs = 0;
		dp.rgvarg = NULL;
		pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dp, NULL, NULL, NULL);
	}

	pDocument->Release();
	pDocuments->Release();
	pWordApp->Release();
	auto end = std::chrono::high_resolution_clock::now();
	auto duration = std::chrono::duration_cast<std::chrono::microseconds>(end - start);

	std::cout << "函数执行时间: " << duration.count() << " 微秒" << std::endl;
	return 0;
}