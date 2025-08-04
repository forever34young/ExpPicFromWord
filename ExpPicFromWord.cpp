#include <comdef.h>
#include <gdiplus.h>
#include <shlwapi.h>
#include <windows.h>

#include <chrono>
#include <iostream>
#include <string>

#pragma comment(lib, "oleaut32.lib")
#pragma comment(lib, "shlwapi.lib")
#pragma comment(lib, "gdiplus.lib")

// Word Application CLSID 和 IID
const CLSID CLSID_WordApplication = {
    0x000209FF,
    0x0000,
    0x0000,
    {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
const CLSID CLSID_WpsApplication = {
    0x000209FF,
    0x0000,
    0x0000,
    {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};
const IID IID__Application = {0x00020970,
                              0x0000,
                              0x0000,
                              {0xC0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46}};

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

DISPPARAMS dpEmpty = {NULL, NULL, 0, 0}; // 空参数
/********************************辅助函数****************************************/
auto getProp = [&](IDispatch *obj, const OLECHAR *name,
                   VARIANT *pResult) -> bool {
  DISPID id;
  if (FAILED(obj->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR *>(&name), 1,
                                LOCALE_USER_DEFAULT, &id)))
    return false;
  return SUCCEEDED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_PROPERTYGET, &dpEmpty, pResult, nullptr,
                               nullptr));
};

bool IsProgIDRegistered(LPCWSTR progID) {
  // 检查 HKCR\<progID>\CLSID 是否存在
  wchar_t keyPath[256];
  wsprintfW(keyPath, L"%s\\CLSID", progID);
  HKEY hKey = nullptr;
  LONG rc = RegOpenKeyExW(HKEY_CLASSES_ROOT, keyPath, 0, KEY_READ, &hKey);
  if (hKey)
    RegCloseKey(hKey);
  return (rc == ERROR_SUCCESS);
}
/********************************！辅助函数***************************************/

/*****************************使用剪切板保存**************************************************/
int GetEncoderClsid(const wchar_t *format, CLSID *pClsid) {
  UINT num = 0, size = 0;
  GetImageEncodersSize(&num, &size);
  if (size == 0)
    return -1;
  auto *pImageCodecInfo = (ImageCodecInfo *)(malloc(size));
  if (!pImageCodecInfo)
    return -1;
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

bool SaveEmfToBitmap(HENHMETAFILE hEmf, const std::wstring &filename,
                     float dpiX = 300.0f, float dpiY = 300.0f) {
  bool success = false;

  Metafile metafile(hEmf);
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

bool SaveImageFromClipboard(const std::wstring &filename) {
  GdiplusStartupInput gdiplusStartupInput;
  ULONG_PTR gdiplusToken;
  GdiplusStartup(&gdiplusToken, &gdiplusStartupInput, nullptr);

  if (!OpenClipboard(nullptr))
    return false;

  bool success = false;

  // 1. 尝试 CF_BITMAP
  HANDLE hClipboardData = GetClipboardData(CF_BITMAP);
  if (hClipboardData) {
    Bitmap *pBitmap = new Bitmap((HBITMAP)hClipboardData, nullptr);
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
      BITMAPINFOHEADER *pBmpInfoHeader =
          (BITMAPINFOHEADER *)GlobalLock(hClipboardData);
      if (pBmpInfoHeader) {
        int width = pBmpInfoHeader->biWidth;
        int height = abs(pBmpInfoHeader->biHeight); // 高度可能是负的
        int rowStride = ((width * pBmpInfoHeader->biBitCount + 31) / 32) * 4;
        BYTE *pixels = (BYTE *)pBmpInfoHeader + pBmpInfoHeader->biSize;
        Bitmap *pBitmap =
            new Bitmap(width, height, rowStride, PixelFormat32bppARGB, pixels);
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
/*****************************!使用剪切板保存**************************************************/

// 使用 EnhMetaFileBits 图元提取图片
int gImageCount = 0;

bool ExtractImageViaEMF(IDispatch *pRange, const std::wstring &filename) {
  // 目标尺寸边界
  const int MAX_WIDTH = 640;
  const int MAX_HEIGHT = 320;

  // 获取EnhMetaFileBits属性ID
  DISPID dispid = 0;
  LPOLESTR emfName = _wcsdup(L"EnhMetaFileBits");
  HRESULT hr = pRange->GetIDsOfNames(IID_NULL, &emfName, 1, LOCALE_USER_DEFAULT,
                                     &dispid);
  free(emfName);
  if (FAILED(hr))
    return false;

  // 调用属性获取方法
  DISPPARAMS dp = {nullptr, nullptr, 0, 0};
  VARIANT result;
  VariantInit(&result);
  hr = pRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                      DISPATCH_PROPERTYGET, &dp, &result, nullptr, nullptr);
  if (FAILED(hr) || result.vt != (VT_ARRAY | VT_UI1)) {
    VariantClear(&result);
    return false;
  }

  // 访问安全数组数据
  SAFEARRAY *psa = result.parray;
  BYTE *pData = nullptr;
  if (FAILED(SafeArrayAccessData(psa, (void **)&pData))) {
    VariantClear(&result);
    return false;
  }

  // 创建增强型图元文件
  LONG lBound = 0, uBound = 0;
  SafeArrayGetLBound(psa, 1, &lBound);
  SafeArrayGetUBound(psa, 1, &uBound);
  DWORD size = uBound - lBound + 1;
  HENHMETAFILE hEmf = SetEnhMetaFileBits(size, pData);
  SafeArrayUnaccessData(psa);
  VariantClear(&result);
  if (!hEmf)
    return false;

  bool success = false;
  {
    // 在 GDI+ 中用 Metafile 加载
    Metafile metafile(hEmf, TRUE /*keep EMF alive*/);

    // 获取 EMF 中记录的实际内容边界（像素坐标）
    Unit unit;
    RectF boundsF;
    Status status = metafile.GetBounds(&boundsF, &unit);
    if (status != Ok) {
      DeleteEnhMetaFile(hEmf);
      gImageCount--;
      return true;
    }

    int x = static_cast<int>(floor(boundsF.X));
    int y = static_cast<int>(floor(boundsF.Y));
    int origWidth = static_cast<int>(ceil(boundsF.Width));
    int origHeight = static_cast<int>(ceil(boundsF.Height));
    if (origWidth <= 0 || origHeight <= 0) {
      DeleteEnhMetaFile(hEmf);
      gImageCount--;
      return true;
    }

    // 计算是否需要缩放及缩放比例
    bool needScaling = (origWidth > MAX_WIDTH || origHeight > MAX_HEIGHT);
    float scale = 1.0f;
    int workingWidth = origWidth;
    int workingHeight = origHeight;

    if (needScaling) {
      // 计算宽和高各自的缩放比例
      float scaleForWidth = static_cast<float>(MAX_WIDTH) / origWidth;
      float scaleForHeight = static_cast<float>(MAX_HEIGHT) / origHeight;

      // 选择较小的缩放比例（确保最长边不超过目标尺寸）
      scale = min(scaleForWidth, scaleForHeight);

      // 计算缩放后的工作尺寸
      workingWidth = static_cast<int>(origWidth * scale);
      workingHeight = static_cast<int>(origHeight * scale);
    }

    // 创建工作位图（原始尺寸或缩放后的尺寸）
    Bitmap workingBmp(workingWidth, workingHeight, PixelFormat32bppARGB);
    workingBmp.SetResolution(96, 96);
    Graphics workingG(&workingBmp);

    // 绘制图像（原始尺寸或缩放后）
    workingG.DrawImage(&metafile, Rect(0, 0, workingWidth, workingHeight), x, y,
                       origWidth, origHeight, UnitPixel);

    // 扫描透明像素区域（在工作位图上进行，已缩放或原始尺寸）
    BitmapData bmpData;
    Rect lockRect(0, 0, workingWidth, workingHeight);
    if (workingBmp.LockBits(&lockRect, ImageLockModeRead, PixelFormat32bppARGB,
                            &bmpData) != Ok) {
      DeleteEnhMetaFile(hEmf);
      return false;
    }

    BYTE *scan0 = static_cast<BYTE *>(bmpData.Scan0);
    int stride = bmpData.Stride;

    int left = workingWidth, right = -1, top = workingHeight, bottom = -1;
    for (int yy = 0; yy < workingHeight; ++yy) {
      BYTE *row = scan0 + yy * stride;
      for (int xx = 0; xx < workingWidth; ++xx) {
        BYTE alpha = row[xx * 4 + 3]; // Alpha 通道
        if (alpha != 0) {
          left = min(left, xx);
          right = max(right, xx);
          top = min(top, yy);
          bottom = max(bottom, yy);
        }
      }
    }

    workingBmp.UnlockBits(&bmpData);

    // 检查裁剪区域有效性
    if (right < left || bottom < top) {
      DeleteEnhMetaFile(hEmf);
      gImageCount--;
      return true;
    }

    // 计算裁剪后的尺寸
    int cropW = right - left + 1;
    int cropH = bottom - top + 1;

    // 剔除过小图像
    if (cropW < 10 || cropH < 10) {
      DeleteEnhMetaFile(hEmf);
      gImageCount--;
      return true;
    }

    // 创建最终位图（裁剪后的尺寸）
    Bitmap outBmp(cropW, cropH, PixelFormat32bppARGB);
    outBmp.SetResolution(96, 96);
    Graphics g2(&outBmp);
    g2.Clear(Color::White); // 清白底

    // 绘制裁剪后的图像
    g2.DrawImage(&workingBmp, Rect(0, 0, cropW, cropH), left, top, cropW, cropH,
                 UnitPixel);

    // 保存图片并应用压缩
    CLSID pngClsid;
    if (GetEncoderClsid(L"image/png", &pngClsid) != -1) {
      EncoderParameters encoderParams;
      encoderParams.Count = 1;
      encoderParams.Parameter[0].Guid = EncoderCompression;
      encoderParams.Parameter[0].Type = EncoderParameterValueTypeLong;
      encoderParams.Parameter[0].NumberOfValues = 1;

      ULONG compressionLevel = 6; // 适中的压缩级别
      encoderParams.Parameter[0].Value = &compressionLevel;

      if (outBmp.Save(filename.c_str(), &pngClsid, &encoderParams) == Ok) {
        success = true;
      }
    }
  }

  DeleteEnhMetaFile(hEmf);
  return success;
}

/****************************************保存嵌入式图片*****************************************************/
// 保存嵌入式图片
bool TryExtractImageFromInlineShape(IDispatch *pInlineShape,
                                    const std::wstring &filename);

int ExportpInlineShapes(IDispatch *pDocument, DISPID dispid,
                        std::wstring outputDir) {
  DISPPARAMS dp = {NULL, NULL, 0, 0};
  // 获取所有内联形状（包括图片）
  HRESULT hr;
  IDispatch *pInlineShapes = NULL;
  OLECHAR *propertyName = _wcsdup(L"InlineShapes");
  hr = pDocument->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                                &dispid);
  if (FAILED(hr)) {
    pDocument->Release();
    std::cout << "无法获取InlineShapes属性" << std::endl;
    return false;
  }
  VARIANT result;
  VariantInit(&result);

  hr = pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_PROPERTYGET, &dpEmpty, &result, NULL, NULL);
  if (FAILED(hr)) {
    std::cout << "无法获取InlineShapes集合" << std::endl;
    return false;
  }
  pInlineShapes = result.pdispVal;

  // 获取形状数量
  propertyName = _wcsdup(L"Count");
  hr = pInlineShapes->GetIDsOfNames(IID_NULL, &propertyName, 1,
                                    LOCALE_USER_DEFAULT, &dispid);
  if (FAILED(hr)) {
    pInlineShapes->Release();
    std::cout << "无法获取Count属性" << std::endl;
    return false;
  }

  long count = 0;
  VariantInit(&result);
  hr = pInlineShapes->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                             DISPATCH_PROPERTYGET, &dpEmpty, &result, NULL,
                             NULL);
  if (SUCCEEDED(hr)) {
    count = result.lVal;
  }

  // 遍历所有形状
  propertyName = _wcsdup(L"Item");
  hr = pInlineShapes->GetIDsOfNames(IID_NULL, &propertyName, 1,
                                    LOCALE_USER_DEFAULT, &dispid);
  if (FAILED(hr)) {
    pInlineShapes->Release();
    std::cout << "无法获取Item方法" << std::endl;
    return false;
  }
  VariantInit(&result);

  static int try_count;                 // 尝试次数
  int imagePageIndex = 1, lastPage = 0; // 上次页码  和本页图片
  for (long i = 1; i <= count; i++) {
    gImageCount = i; // 计数
    try_count = 0;
    auto dispid1 = dispid;
    VARIANT varIndex;
    varIndex.vt = VT_I4;
    varIndex.lVal = i;

    dp.cArgs = 1;
    dp.rgvarg = &varIndex;

    VariantInit(&result);
    hr = pInlineShapes->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &dp, &result, NULL, NULL);
    if (FAILED(hr) || result.pdispVal == NULL) {
      _com_error err(hr);
      std::cout << "获取第 " << i << " 个形状失败 (HRESULT: 0x" << std::hex
                << hr << "): " << err.ErrorMessage() << std::endl;
      VariantClear(&result);
      continue;
    }

    IDispatch *pInlineShape = result.pdispVal;

  gotry:
    IDispatch *pRange = NULL;
    propertyName = _wcsdup(L"Range");
    hr = pInlineShape->GetIDsOfNames(IID_NULL, &propertyName, 1,
                                     LOCALE_USER_DEFAULT, &dispid1);
    if (FAILED(hr)) {
      pInlineShape->Release();
      continue;
    }

    VariantInit(&result);
    dp.cArgs = 0;
    dp.rgvarg = NULL;
    hr = pInlineShape->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
    if (FAILED(hr) || result.pdispVal == NULL) {
      pInlineShape->Release();
      continue;
    }
    pRange = result.pdispVal;

    // 获取页码
    long pageNumber = 0;
    propertyName = _wcsdup(L"Information");
    hr = pRange->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                               &dispid1);
    if (SUCCEEDED(hr)) {
      VARIANT varWdInformation;
      varWdInformation.vt = VT_I4;
      varWdInformation.lVal = 3; // wdActiveEndPageNumber

      dp.cArgs = 1;
      dp.rgvarg = &varWdInformation;

      VariantInit(&result);
      hr = pRange->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
      if (SUCCEEDED(hr)) {
        pageNumber = result.lVal;
      } else {
        std::cout << "get pageNumber failed 1" << std::endl;
        hr = pRange->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT,
                            DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
        if (SUCCEEDED(hr)) {
          pageNumber = result.lVal;
        } else {
          std::cout << "get pageNumber failed 2" << std::endl;
          hr = pRange->Invoke(dispid1, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
          if (SUCCEEDED(hr)) {
            pageNumber = result.lVal;
          } else {
            std::cout << "get pageNumber failed 3" << std::endl;
          }
        }
      }
    }
    if (pageNumber == 0) {
      std::cout << "get pageNumber failed" << std::endl;
      int stop_here = 0;
    }
    if (pageNumber == lastPage) {
      imagePageIndex++;
    } else {
      imagePageIndex = 1;
    }

    // 创建输出文件名
    wchar_t filename[MAX_PATH];
    swprintf_s(filename, L"%s\\image_%d_%d.png", outputDir.c_str(), pageNumber,
               imagePageIndex);

    lastPage = pageNumber;

    // 提取并保存图片
    if (TryExtractImageFromInlineShape(pInlineShape, filename)) {
      std::cout << i << "已保存图片 " << std::endl;
    } else {
      std::cout << i << "无法保存图片" << std::endl;
      if (pageNumber == lastPage) {
        imagePageIndex--;
      }
      try_count++;
      if (try_count <= 3) {
        goto gotry;
      } else {
        std::cout << i
                  << "--------------------无法保存图片----------------------"
                  << std::endl;
      }
    }
    pRange->Release();
    // 清理资源
    pInlineShape->Release();
    VariantClear(&result);
  }
  return 0;
}

// 保存嵌入式图片
bool TryExtractImageFromInlineShape(IDispatch *pInlineShape,
                                    const std::wstring &filename) {
  DISPID dispid;

  VARIANT result;
  VariantInit(&result);
  DISPPARAMS dp = {NULL, NULL, 0, 0};

  OLECHAR *propertyName = _wcsdup(L"Range");
  HRESULT hr = pInlineShape->GetIDsOfNames(IID_NULL, &propertyName, 1,
                                           LOCALE_USER_DEFAULT, &dispid);
  free(propertyName);
  if (FAILED(hr)) {
    return false;
  }

  hr = pInlineShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                            DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
  if (FAILED(hr) || result.vt != VT_DISPATCH) {
    VariantClear(&result);
    return false;
  }
  IDispatch *pRange = result.pdispVal;

  if (ExtractImageViaEMF(pRange, filename)) {
    // std::cout << "[EMF] 成功提取图像：" << std::endl;
    return true;
  }

  // 如果EMF提取失败，尝试使用CopyAsPicture方法
  propertyName = _wcsdup(L"CopyAsPicture");
  hr = pRange->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                             &dispid);

  if (SUCCEEDED(hr)) {
    if (OpenClipboard(nullptr)) {
      EmptyClipboard();
      CloseClipboard();
      hr = pRange->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_METHOD, &dp, NULL, NULL, NULL);
      Sleep(10);

      if (SaveImageFromClipboard(filename)) {
        return true;
      }
    }
  }
  pRange->Release();
  free(propertyName);

  return false;
}

/****************************************！保存嵌入式图片*****************************************************/

/****************************************保存页眉页脚图片*****************************************************/
// 获取节的起始页码
long GetSectionStartPage(IDispatch *pSection);
// 判断是不是在表格中
bool IsRangeInsideTable(IDispatch *pRange);
// 清理页码
void DeletePageNumberFieldsOnly(IDispatch *pHF);
//  清理文字
void DeleteNonTableParagraphsInRange(IDispatch *pRange);

int last_section = 1;
void ExtractImagesFromHeaders(IDispatch *pDocument, IDispatch *pWordApp,
                              const std::wstring &outputDir) {
  DISPID dispid;
  VARIANT result;

  VARIANT sectionsVar;

  DISPPARAMS dp = {NULL, NULL, 0, 0};
  VariantInit(&sectionsVar);
  if (!getProp(pDocument, L"Sections", &sectionsVar) ||
      sectionsVar.vt != VT_DISPATCH)
    return;
  IDispatch *pSections = sectionsVar.pdispVal;

  DISPID itemDispid;
  OLECHAR *itemName = _wcsdup(L"Item");
  pSections->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT,
                           &itemDispid);

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
    pSections->Invoke(itemDispid, IID_NULL, LOCALE_USER_DEFAULT,
                      DISPATCH_METHOD, &dp, &result, nullptr, nullptr);
    if (!result.pdispVal)
      continue;
    IDispatch *pSection = result.pdispVal;
    long startPage = GetSectionStartPage(pSection);

    for (const OLECHAR *type :
         {_wcsdup(L"Headers") /*, _wcsdup(L"Footers")*/}) {
      VARIANT hfVar;
      VariantInit(&hfVar);

      dp = {nullptr, nullptr, 0, 0};
      if (!getProp(pSection, type, &hfVar) || hfVar.vt != VT_DISPATCH)
        continue;
      IDispatch *pCollection = hfVar.pdispVal;

      getProp(pCollection, L"Count", &countVar);
      long hfCount = countVar.lVal;

      DISPID hfItemDispid;
      pCollection->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT,
                                 &hfItemDispid);

      for (long j = 1; j <= 1; ++j) {
        index.lVal = j;
        dp.rgvarg = &index;
        dp.cArgs = 1;
        VariantInit(&result);
        pCollection->Invoke(hfItemDispid, IID_NULL, LOCALE_USER_DEFAULT,
                            DISPATCH_METHOD, &dp, &result, nullptr, nullptr);
        if (!result.pdispVal)
          continue;
        IDispatch *pHF = result.pdispVal;

        VARIANT rangeVar;
        VariantInit(&rangeVar);
        dp = {nullptr, nullptr, 0, 0};
        if (!getProp(pHF, L"Range", &rangeVar) || rangeVar.vt != VT_DISPATCH)
          continue;
        IDispatch *pRange = rangeVar.pdispVal;

        DeletePageNumberFieldsOnly(pHF);
        /**********************************************************/

        wchar_t filename[256];
        swprintf_s(filename, L"%s\\header_%d_%d.png", outputDir.c_str(),
                   last_section, imageIndex++);
        if (ExtractImageViaEMF(pRange, filename)) {
        }

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

// 获取节的起始页码
long GetSectionStartPage(IDispatch *pSection) {
  VARIANT result;
  VariantInit(&result);

  auto getProp = [&](IDispatch *obj, const OLECHAR *name,
                     VARIANT *pResult) -> bool {
    DISPPARAMS dp = {nullptr, nullptr, 0, 0};
    DISPID id;
    if (FAILED(obj->GetIDsOfNames(IID_NULL, const_cast<LPOLESTR *>(&name), 1,
                                  LOCALE_USER_DEFAULT, &id)))
      return false;
    return SUCCEEDED(obj->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &dp, pResult, nullptr,
                                 nullptr));
  };

  VARIANT rangeVar;
  VariantInit(&rangeVar);
  if (!getProp(pSection, L"Range", &rangeVar) || rangeVar.vt != VT_DISPATCH)
    return -1;
  IDispatch *pRange = rangeVar.pdispVal;

  DISPID dispidInfo;
  OLECHAR *infoProp = _wcsdup(L"Information");
  VARIANT arg;
  arg.vt = VT_I4;
  arg.lVal = 3;

  if (SUCCEEDED(pRange->GetIDsOfNames(IID_NULL, &infoProp, 1,
                                      LOCALE_USER_DEFAULT, &dispidInfo))) {
    DISPPARAMS dpInfo = {&arg, nullptr, 1, 0};
    if (SUCCEEDED(pRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_PROPERTYGET, &dpInfo, &result,
                                 nullptr, nullptr))) {
      pRange->Release();
      return result.lVal;
    }
  }
  pRange->Release();
  return -1;
}

// 判断是不是在表格中
bool IsRangeInsideTable(IDispatch *pRange) {
  if (!pRange)
    return false;
  DISPID dispidInfo;
  OLECHAR *method = _wcsdup(L"Information");
  if (FAILED(pRange->GetIDsOfNames(IID_NULL, &method, 1, LOCALE_USER_DEFAULT,
                                   &dispidInfo)))
    return false;

  VARIANT arg;
  arg.vt = VT_I4;
  arg.lVal = 12; // wdWithInTable
  DISPPARAMS dpInfo = {&arg, nullptr, 1, 0};
  VARIANT result;
  VariantInit(&result);
  if (SUCCEEDED(pRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT,
                               DISPATCH_METHOD, &dpInfo, &result, nullptr,
                               nullptr))) {
    return (result.vt == VT_BOOL && result.boolVal == VARIANT_TRUE);
  }
  return false;
}

// 辅助函数：删除指定Range中的所有页码字段
bool HasFieldsInRange(IDispatch *pRange) {
  bool has_fields = false;
  VARIANT fieldsVar;
  VariantInit(&fieldsVar);
  if (getProp(pRange, L"Fields", &fieldsVar) && fieldsVar.vt == VT_DISPATCH) {
    IDispatch *pFields = fieldsVar.pdispVal;
    VARIANT countVar;
    if (getProp(pFields, L"Count", &countVar) && countVar.vt == VT_I4) {
      long count = countVar.lVal;
      DISPID dispidItem;
      OLECHAR *itemName = _wcsdup(L"Item");
      pFields->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT,
                             &dispidItem);
      free(itemName);

      // 注意：必须从后向前删除
      for (long i = count; i >= 1; --i) {
        VARIANT index;
        index.vt = VT_I4;
        index.lVal = i;
        DISPPARAMS dpItem = {&index, nullptr, 1, 0};
        VARIANT fieldVar;
        VariantInit(&fieldVar);
        if (SUCCEEDED(pFields->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                                      DISPATCH_METHOD, &dpItem, &fieldVar,
                                      nullptr, nullptr))) {
          if (fieldVar.vt == VT_DISPATCH && fieldVar.pdispVal) {
            VARIANT typeVar;
            if (getProp(fieldVar.pdispVal, L"Type", &typeVar) &&
                typeVar.vt == VT_I4) {
              // 页码字段类型:
              // wdFieldPage(13)/wdFieldPageRef(26)/wdFieldSection(33)
              if (typeVar.lVal == 13 || typeVar.lVal == 26 ||
                  typeVar.lVal == 33) {

                has_fields = true;
              }
            }
            fieldVar.pdispVal->Release();
          }
        }
      }
    }
    pFields->Release();
  }
  return has_fields;
}

// 递归处理Range中的所有形状（特别是文本框）
void ProcessShapesInRange(IDispatch *pRange) {
  VARIANT shapesVar;
  VariantInit(&shapesVar);

  // 获取当前Range中的所有形状
  if (getProp(pRange, L"Shapes", &shapesVar) && shapesVar.vt == VT_DISPATCH) {
    IDispatch *pShapes = shapesVar.pdispVal;
    VARIANT countVar;

    if (getProp(pShapes, L"Count", &countVar) && countVar.vt == VT_I4) {
      long shapeCount = countVar.lVal;

      for (long i = shapeCount; i >= 1; --i) {
        // 获取第i个形状
        VARIANT index;
        index.vt = VT_I4;
        index.lVal = i;
        DISPPARAMS dp = {&index, nullptr, 1, 0};

        DISPID dispidItem;
        OLECHAR *itemName = const_cast<OLECHAR *>(L"Item");
        pShapes->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT,
                               &dispidItem);

        VARIANT shapeVar;
        VariantInit(&shapeVar);
        if (SUCCEEDED(pShapes->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                                      DISPATCH_METHOD, &dp, &shapeVar, nullptr,
                                      nullptr))) {
          if (shapeVar.vt == VT_DISPATCH && shapeVar.pdispVal) {
            IDispatch *pShape = shapeVar.pdispVal;

            // 尝试获取 TextFrame
            VARIANT textFrameVar;
            if (getProp(pShape, L"TextFrame", &textFrameVar) &&
                textFrameVar.vt == VT_DISPATCH) {
              IDispatch *pTextFrame = textFrameVar.pdispVal;

              // 尝试获取 TextRange
              VARIANT textRangeVar;
              if (getProp(pTextFrame, L"TextRange", &textRangeVar) &&
                  textRangeVar.vt == VT_DISPATCH) {
                IDispatch *pTextRange = textRangeVar.pdispVal;

                // 判断是否包含字段
                bool has_fields = HasFieldsInRange(pTextRange);

                // 判断是否包含表格
                bool has_table = false;
                VARIANT tablesVar;
                if (getProp(pTextRange, L"Tables", &tablesVar) &&
                    tablesVar.vt == VT_DISPATCH) {
                  IDispatch *pTables = tablesVar.pdispVal;
                  VARIANT tableCountVar;
                  if (getProp(pTables, L"Count", &tableCountVar) &&
                      tableCountVar.vt == VT_I4) {
                    has_table = (tableCountVar.lVal > 0);
                  }
                  pTables->Release();
                }

                // 满足：有字段，且无表格 → 删除整个 Shape
                if (has_fields && !has_table) {
                  DISPID dispidDelete;
                  OLECHAR *deleteName = const_cast<OLECHAR *>(L"Delete");
                  pShape->GetIDsOfNames(IID_NULL, &deleteName, 1,
                                        LOCALE_USER_DEFAULT, &dispidDelete);
                  DISPPARAMS dpDel = {nullptr, nullptr, 0, 0};
                  pShape->Invoke(dispidDelete, IID_NULL, LOCALE_USER_DEFAULT,
                                 DISPATCH_METHOD, &dpDel, nullptr, nullptr,
                                 nullptr);
                } else {
                  // 否则递归处理嵌套 Range 中的 Shape
                  ProcessShapesInRange(pTextRange);
                }

                pTextRange->Release();
              }

              pTextFrame->Release();
            }

            pShape->Release();
          }
        }
      }
    }
    pShapes->Release();
  }
}

// 清理页码
void DeletePageNumberFieldsOnly(IDispatch *pHF) {
  if (!pHF)
    return;

  ProcessShapesInRange(pHF);
  //// ===== 删除 Range 中的页码字段 =====
  // VARIANT rangeVar;
  // VariantInit(&rangeVar);
  // if (!getProp(pHF, L"Range", &rangeVar) || rangeVar.vt != VT_DISPATCH)
  //   return;
  // IDispatch *pRange = rangeVar.pdispVal;

  // VARIANT fieldsVar;
  // VariantInit(&fieldsVar);
  // if (getProp(pRange, L"Fields", &fieldsVar) && fieldsVar.vt == VT_DISPATCH)
  // {
  //   IDispatch *pFields = fieldsVar.pdispVal;
  //   VARIANT countVar;
  //   if (getProp(pFields, L"Count", &countVar) && countVar.vt == VT_I4) {
  //     long count = countVar.lVal;
  //     DISPID dispidItem, dispidDelete;
  //     OLECHAR *itemName = _wcsdup(L"Item");
  //     OLECHAR *deleteName = _wcsdup(L"Delete");
  //     pFields->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT,
  //                            &dispidItem);
  //     free(itemName);
  //     for (long i = count; i >= 1; --i) {
  //       VARIANT index;
  //       index.vt = VT_I4;
  //       index.lVal = i;
  //       DISPPARAMS dpItem = {&index, nullptr, 1, 0};
  //       VARIANT fieldVar;
  //       VariantInit(&fieldVar);
  //       if (SUCCEEDED(pFields->Invoke(dispidItem, IID_NULL,
  //       LOCALE_USER_DEFAULT,
  //                                     DISPATCH_METHOD, &dpItem, &fieldVar,
  //                                     nullptr, nullptr))) {
  //         if (fieldVar.vt == VT_DISPATCH && fieldVar.pdispVal) {
  //           VARIANT typeVar;
  //           if (getProp(fieldVar.pdispVal, L"Type", &typeVar) &&
  //               typeVar.vt == VT_I4) {
  //             if (typeVar.lVal == 13 || typeVar.lVal == 26 ||
  //                 typeVar.lVal == 33) { // 13: wdFieldPage, 26:
  //                 wdFieldPageRef,
  //                                       // 33: wdFieldSection
  //               DISPPARAMS dpDel = {nullptr, nullptr, 0, 0};
  //               pRange->GetIDsOfNames(IID_NULL, &deleteName, 1,
  //                                     LOCALE_USER_DEFAULT, &dispidDelete);
  //               auto hr = pRange->Invoke(dispidDelete, IID_NULL,
  //                                        LOCALE_USER_DEFAULT,
  //                                        DISPATCH_METHOD, &dpDel, nullptr,
  //                                        nullptr, nullptr);
  //             }
  //           }
  //           fieldVar.pdispVal->Release();
  //         }
  //       }
  //     }
  //     free(deleteName);
  //   }
  //   pFields->Release();
  // }
  // DeleteNonTableParagraphsInRange(pRange);
  // pRange->Release();
}

//  清理文字
void DeleteNonTableParagraphsInRange(IDispatch *pRange) {
  if (!pRange)
    return;

  VARIANT parasVar;
  if (!getProp(pRange, L"Paragraphs", &parasVar) || parasVar.vt != VT_DISPATCH)
    return;
  IDispatch *pParas = parasVar.pdispVal;

  VARIANT countVar;
  if (!getProp(pParas, L"Count", &countVar) || countVar.vt != VT_I4) {
    pParas->Release();
    return;
  }
  long pCount = countVar.lVal;

  DISPID dispidParaItem;
  OLECHAR *name = _wcsdup(L"Item");
  pParas->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT,
                        &dispidParaItem);

  for (long i = 1; i <= pCount; ++i) {
    VARIANT pidx = {VT_I4};
    pidx.lVal = i;
    DISPPARAMS dpPara = {&pidx, nullptr, 1, 0};
    VARIANT paraVar;
    if (FAILED(pParas->Invoke(dispidParaItem, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_METHOD, &dpPara, &paraVar, nullptr,
                              nullptr)))
      continue;
    if (paraVar.vt != VT_DISPATCH || !paraVar.pdispVal)
      continue;

    IDispatch *pPara = paraVar.pdispVal;

    VARIANT prRangeVar;
    if (getProp(pPara, L"Range", &prRangeVar) && prRangeVar.vt == VT_DISPATCH) {
      IDispatch *pPR = prRangeVar.pdispVal;

      bool clear = true;
      if (1 && IsRangeInsideTable(pPR)) {
        clear = false;
      }

      if (clear) {
        // 遍历 Range.Characters
        VARIANT charsVar;
        if (getProp(pPR, L"Characters", &charsVar) &&
            charsVar.vt == VT_DISPATCH) {
          IDispatch *pChars = charsVar.pdispVal;

          VARIANT charCountVar;
          if (getProp(pChars, L"Count", &charCountVar) &&
              charCountVar.vt == VT_I4) {
            long charCount = charCountVar.lVal;

            DISPID dispidCharItem, dispidDelete;
            OLECHAR *itemName = _wcsdup(L"Item");
            OLECHAR *delName = _wcsdup(L"Delete");
            pChars->GetIDsOfNames(IID_NULL, &itemName, 1, LOCALE_USER_DEFAULT,
                                  &dispidCharItem);

            for (long c = charCount; c >= 1; --c) {
              VARIANT idx;
              idx.vt = VT_I4;
              idx.lVal = c;
              DISPPARAMS dpChar = {&idx, nullptr, 1, 0};
              VARIANT charVar;
              if (SUCCEEDED(pChars->Invoke(
                      dispidCharItem, IID_NULL, LOCALE_USER_DEFAULT,
                      DISPATCH_METHOD, &dpChar, &charVar, nullptr, nullptr))) {
                if (charVar.vt == VT_DISPATCH && charVar.pdispVal) {
                  IDispatch *pCharRange = charVar.pdispVal;

                  // 获取 .Text
                  VARIANT tVar;
                  if (getProp(pCharRange, L"Text", &tVar) &&
                      tVar.vt == VT_BSTR && tVar.bstrVal) {
                    wchar_t ch = tVar.bstrVal[0];
                    // 仅删除普通字符，跳过回车符（\r）、图片占位符等
                    if (ch >= 32 && ch != 0x0D && ch != 0x07 && ch != 0x01 &&
                        ch != 0x02 && ch != 0x2F) {
                      if (SUCCEEDED(pCharRange->GetIDsOfNames(
                              IID_NULL, &delName, 1, LOCALE_USER_DEFAULT,
                              &dispidDelete))) {
                        DISPPARAMS dpDel = {nullptr, nullptr, 0, 0};
                        pCharRange->Invoke(dispidDelete, IID_NULL,
                                           LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                                           &dpDel, nullptr, nullptr, nullptr);
                      }
                    }
                  }
                  pCharRange->Release();
                }
              }
            }
          }
          pChars->Release();
        }
      }
      pPR->Release();
    }
    pPara->Release();
  }

  pParas->Release();
}
/****************************************!保存页眉页脚图片*****************************************************/

/****************************************保存浮动图片*****************************************************/
// 多种方式尝试保存浮动图片
bool TryExtractImageSmart(IDispatch *pWordApp, IDispatch *pShape,
                          const std::wstring &filename);
// 保存浮动类型的图片
bool TryExtractImageFromShape(IDispatch *pWordApp, IDispatch *pShape,
                              const std::wstring &filename);

int ExportFloatingShapes(IDispatch *pDocument, IDispatch *pWordApp,
                         const std::wstring &outputDir) {
  HRESULT hr;
  DISPID dispid;
  DISPPARAMS noArgs = {NULL, NULL, 0, 0};
  VARIANT result;
  IDispatch *pShapes = nullptr;

  // 获取 Shapes
  OLECHAR *propShapes = _wcsdup(L"Shapes");
  hr = pDocument->GetIDsOfNames(IID_NULL, &propShapes, 1, LOCALE_USER_DEFAULT,
                                &dispid);
  if (FAILED(hr))
    return -1;

  hr = pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
  if (FAILED(hr) || result.vt != VT_DISPATCH)
    return -1;
  pShapes = result.pdispVal;

  // 获取 Shapes.Count
  DISPID dispidCount;
  OLECHAR *countProp = _wcsdup(L"Count");
  hr = pShapes->GetIDsOfNames(IID_NULL, &countProp, 1, LOCALE_USER_DEFAULT,
                              &dispidCount);
  if (FAILED(hr))
    return -1;

  VariantInit(&result);
  hr = pShapes->Invoke(dispidCount, IID_NULL, LOCALE_USER_DEFAULT,
                       DISPATCH_PROPERTYGET, &noArgs, &result, NULL, NULL);
  if (FAILED(hr))
    return -1;
  long shapeCount = result.lVal;

  // 获取 Shapes.Item 方法
  DISPID dispidItem;
  OLECHAR *itemMethod = _wcsdup(L"Item");
  hr = pShapes->GetIDsOfNames(IID_NULL, &itemMethod, 1, LOCALE_USER_DEFAULT,
                              &dispidItem);
  if (FAILED(hr))
    return -1;

  static int float_try_count;
  int imagePageIndex = 1, lastPage = 0; // 上次页码  和本页图片
  for (long i = 1; i <= shapeCount; ++i) {
    float_try_count = 0;
  gofloattry:
    VARIANT index;
    index.vt = VT_I4;
    index.lVal = i;
    DISPPARAMS dpItem = {&index, NULL, 1, 0};
    VariantInit(&result);
    hr = pShapes->Invoke(dispidItem, IID_NULL, LOCALE_USER_DEFAULT,
                         DISPATCH_METHOD, &dpItem, &result, NULL, NULL);
    if (FAILED(hr) || result.vt != VT_DISPATCH)
      continue;
    IDispatch *pShape = result.pdispVal;

    // 获取 Anchor.Range.Information(wdActiveEndPageNumber)
    int pageNumber = 0;
    DISPID dispidAnchor, dispidRange, dispidInfo;
    VARIANT anchorVar;
    VariantInit(&anchorVar);
    OLECHAR *anchorProp = _wcsdup(L"Anchor");
    hr = pShape->GetIDsOfNames(IID_NULL, &anchorProp, 1, LOCALE_USER_DEFAULT,
                               &dispidAnchor);
    if (SUCCEEDED(hr)) {
      pShape->Invoke(dispidAnchor, IID_NULL, LOCALE_USER_DEFAULT,
                     DISPATCH_PROPERTYGET, &noArgs, &anchorVar, NULL, NULL);
      if (anchorVar.vt == VT_DISPATCH) {
        IDispatch *pRange = anchorVar.pdispVal;
        OLECHAR *rangeProp = _wcsdup(L"Information");
        hr = pRange->GetIDsOfNames(IID_NULL, &rangeProp, 1, LOCALE_USER_DEFAULT,
                                   &dispidInfo);
        if (SUCCEEDED(hr)) {
          VARIANT arg;
          arg.vt = VT_I4;
          arg.lVal = 3; // wdActiveEndPageNumber
          DISPPARAMS dpInfo = {&arg, NULL, 1, 0};
          VariantInit(&result);
          hr = pRange->Invoke(dispidInfo, IID_NULL, LOCALE_USER_DEFAULT,
                              DISPATCH_PROPERTYGET, &dpInfo, &result, NULL,
                              NULL);
          if (SUCCEEDED(hr))
            pageNumber = result.lVal;
        }
        pRange->Release();
      }
    }
    if (pageNumber == 0) {
      pageNumber = 1;
    }
    if (pageNumber == lastPage) {
      imagePageIndex++;
    } else {
      imagePageIndex = 1;
    }
    // 生成文件名
    wchar_t filename[MAX_PATH];
    swprintf_s(filename, L"%s\\float_image_%d_%d.png", outputDir.c_str(),
               pageNumber, imagePageIndex);

    if (TryExtractImageSmart(pWordApp, pShape, filename)) {
      gImageCount++;
      std::cout << gImageCount << "已保存浮动图片: " << std::endl;
      lastPage = pageNumber;
    } else {
      if (pageNumber == lastPage) {
        imagePageIndex--;
      }
      std::cout << gImageCount << "无法保存图片: page = " << pageNumber
                << std::endl;
      float_try_count++;
      if (float_try_count <= 3) {
        goto gofloattry;
      }
    }
    pShape->Release();
  }

  pShapes->Release();
  return 0;
}

bool TryExtractImageSmart(IDispatch *pWordApp, IDispatch *pShape,
                          const std::wstring &filename) {
  // 尝试方法 1：CopyAsPicture + 剪贴板提取
  if (TryExtractImageFromShape(pWordApp, pShape, filename)) {
    // std::cout << "成功提取图片" << std::endl;
    return true;
  }

  // 尝试方法 2：Export 方法（某些图片/图表/艺术字支持）
  DISPID dispidExport;
  OLECHAR *exportMethod = _wcsdup(L"Export");
  HRESULT hr = pShape->GetIDsOfNames(IID_NULL, &exportMethod, 1,
                                     LOCALE_USER_DEFAULT, &dispidExport);
  free(exportMethod);
  if (SUCCEEDED(hr)) {
    VARIANT varFilename, varFormat;
    varFilename.vt = VT_BSTR;
    varFilename.bstrVal = SysAllocString(filename.c_str());

    varFormat.vt = VT_BSTR;
    varFormat.bstrVal = SysAllocString(L"PNG"); // 你也可以试 EMF

    VARIANT args[2] = {varFormat, varFilename};
    DISPPARAMS dpExport = {args, nullptr, 2, 0};

    hr = pShape->Invoke(dispidExport, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &dpExport, nullptr, nullptr, nullptr);
    VariantClear(&varFilename);
    VariantClear(&varFormat);

    if (SUCCEEDED(hr)) {
      std::cout << "[Smart] 使用 Export 成功导出图片" << std::endl;
      return true;
    }
  }

  // 尝试方法 3：ConvertToInlineShape → TryExtractImageFromInlineShape
  DISPID dispidConvert;
  OLECHAR *convertMethod = _wcsdup(L"ConvertToInlineShape");
  hr = pShape->GetIDsOfNames(IID_NULL, &convertMethod, 1, LOCALE_USER_DEFAULT,
                             &dispidConvert);
  free(convertMethod);

  if (SUCCEEDED(hr)) {
    DISPPARAMS noArgs = {NULL, NULL, 0, 0};
    VARIANT result;
    VariantInit(&result);
    hr = pShape->Invoke(dispidConvert, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_METHOD, &noArgs, &result, NULL, NULL);
    if (SUCCEEDED(hr) && result.vt == VT_DISPATCH && result.pdispVal) {
      IDispatch *pInlineShape = result.pdispVal;
      bool success = false;
      for (size_t i = 0; i < 5; i++) {
        success = TryExtractImageFromInlineShape(pInlineShape, filename);
        if (success) {
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
  return false;
}

//  首先选中，然后保存图片
bool TryExtractImageFromShape(IDispatch *pWordApp, IDispatch *pShape,
                              const std::wstring &filename) {
  HRESULT hr;
  DISPID dispid;

  // 1. 选中图形
  OLECHAR *selectMethod = _wcsdup(L"Select");
  hr = pShape->GetIDsOfNames(IID_NULL, &selectMethod, 1, LOCALE_USER_DEFAULT,
                             &dispid);
  if (FAILED(hr))
    return false;

  DISPPARAMS noArgs = {NULL, NULL, 0, 0};
  hr = pShape->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                      &noArgs, NULL, NULL, NULL);
  if (FAILED(hr))
    return false;

  // 2. 获取 Selection
  IDispatch *pSelection = nullptr;
  VARIANT selResult;
  VariantInit(&selResult);
  DISPID dispidSelection;
  OLECHAR *selProp = _wcsdup(L"Selection");
  hr = pWordApp->GetIDsOfNames(IID_NULL, &selProp, 1, LOCALE_USER_DEFAULT,
                               &dispidSelection);
  if (FAILED(hr))
    return false;

  hr = pWordApp->Invoke(dispidSelection, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_PROPERTYGET, &noArgs, &selResult, NULL, NULL);
  if (FAILED(hr) || selResult.vt != VT_DISPATCH)
    return false;

  pSelection = selResult.pdispVal;

  LPOLESTR emfName = _wcsdup(L"EnhMetaFileBits");
  hr = pSelection->GetIDsOfNames(IID_NULL, &emfName, 1, LOCALE_USER_DEFAULT,
                                 &dispid);
  free(emfName);

  if (SUCCEEDED(hr)) {
    if (ExtractImageViaEMF(pSelection, filename)) {
      // std::cout << "[EMF] 成功提取图像：" << std::endl;
      return true;
    }
  }

  // 3. Selection.Copy()
  OLECHAR *copyMethod = _wcsdup(L"CopyAsPicture");
  hr = pSelection->GetIDsOfNames(IID_NULL, &copyMethod, 1, LOCALE_USER_DEFAULT,
                                 &dispid);
  if (SUCCEEDED(hr)) {
    pSelection->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                       &noArgs, NULL, NULL, NULL);
  }
  pSelection->Release();

  // 4. 从剪贴板提取图像
  Sleep(10); // 等待剪贴板准备
  return SaveImageFromClipboard(filename);
}

/****************************************!保存浮动图片*****************************************************/

// 递归创建宽字符目录 该函数会遍历路径中的每个子目录，如果子目录不存在则创建它
bool CreateDirectoryRecursive(const std::wstring &path) {
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
        // DWORD error = GetLastError();
        // std::wcerr << L"Failed to create directory: " << currentPath << L",
        // Error code: " << error << std::endl;
        return false;
      }
    }
  }
  // 处理最后一个子目录
  if (!path.empty() && !PathFileExistsW(path.c_str())) {
    if (!CreateDirectoryW(path.c_str(), NULL)) {
      // 获取错误代码
      // DWORD error = GetLastError();
      // std::wcerr << L"Failed to create directory: " << path << L", Error
      // code: " << error << std::endl;
      return false;
    }
  }
  return true;
}

int wmain(int argc, wchar_t *argv[]) {
#if 0
  if (argc != 3) {
    return 1;
  }
  std::wstring docPath = argv[1];
  std::wstring outputDir = argv[2];
#else
  // test
  std::wstring docPath = L"F:/1010.docx";
  std::wstring outputDir = L"F:/66pic";
#endif
  auto start = std::chrono::high_resolution_clock::now();
  ComInitializer comInit;

  if (!PathFileExistsW(outputDir.c_str())) {
    if (CreateDirectoryRecursive(outputDir)) {
      std::cout << "Directory created successfully." << std::endl;
    } else {
      std::cout << "Failed to create directory." << std::endl;
    }
  }
  // 创建Word应用程序对象
  IDispatch *pWordApp = NULL;
  bool hasWord = IsProgIDRegistered(L"Word.Application");
  bool hasWPS = IsProgIDRegistered(L"KWPS.Application");

  HRESULT hr = E_FAIL;
  if (hasWPS) {
    CLSID clsid;
    HRESULT hrp = CLSIDFromProgID(L"KWPS.Application", &clsid);
    // 创建 Word 应用程序对象
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                          (void **)&pWordApp);
  }

  if (FAILED(hr) && hasWord) {
    std::cout << "use wps failed" << std::endl;
    CLSID clsid;
    HRESULT hro = CLSIDFromProgID(L"Word.Application", &clsid);
    hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch,
                          (void **)&pWordApp);
  }
  if (FAILED(hr)) {
    throw std::runtime_error(
        "1 Unable to create Wps/Word application instance");
  }

  // 设置Word不可见
  DISPID dispid;
  OLECHAR *propertyName = _wcsdup(L"Visible");
  hr = pWordApp->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                               &dispid);
  if (SUCCEEDED(hr)) {
    VARIANT var;
    var.vt = VT_BOOL;
    var.boolVal = VARIANT_FALSE;
    DISPPARAMS dp = {&var, NULL, 1, 0};
    pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                     DISPATCH_PROPERTYPUT, &dp, NULL, NULL, NULL);
  }

  // 打开文档
  IDispatch *pDocuments = NULL;
  propertyName = _wcsdup(L"Documents");
  hr = pWordApp->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                               &dispid);
  if (FAILED(hr)) {
    pWordApp->Release();
    std::cout << "无法获取Documents属性" << std::endl;
    return false;
  }
  VARIANT result;
  VariantInit(&result);
  DISPPARAMS dp = {NULL, NULL, 0, 0};
  hr = pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                        DISPATCH_PROPERTYGET, &dp, &result, NULL, NULL);
  if (FAILED(hr) || result.pdispVal == NULL) {
    pWordApp->Release();
    std::cout << "无法获取Documents集合" << std::endl;
    return false;
  }
  pDocuments = result.pdispVal;

  // 调用Documents.Open方法
  IDispatch *pDocument = NULL;
  propertyName = _wcsdup(L"Open");
  hr = pDocuments->GetIDsOfNames(IID_NULL, &propertyName, 1,
                                 LOCALE_USER_DEFAULT, &dispid);
  if (FAILED(hr)) {
    pDocuments->Release();
    pWordApp->Release();
    std::cout << "无法获取Open方法" << std::endl;
    return false;
  }

  VARIANT args[1];
  args[0].vt = VT_BSTR;
  args[0].bstrVal = SysAllocString(docPath.c_str());
  DISPPARAMS dpOpen = {args, nullptr, 1, 0};
  VariantInit(&result);
  hr = pDocuments->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT,
                          DISPATCH_METHOD, &dpOpen, &result, nullptr, nullptr);
  pDocument = result.pdispVal;

  if (pDocument == nullptr) {
    std::cout << "Open Document failed" << std::endl;
    return -1;
  }
  std::cout << "Open Document success" << std::endl;

  ExportpInlineShapes(pDocument, dispid, outputDir);

  std::cout << "====================================" << std::endl;
  ExtractImagesFromHeaders(pDocument, pWordApp, outputDir);
  std::cout << "====================================" << std::endl;
  ExportFloatingShapes(pDocument, pWordApp, outputDir);

  // 关闭文档
  propertyName = _wcsdup(L"Close");
  hr = pDocument->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                                &dispid);
  if (SUCCEEDED(hr)) {
    dp.cArgs = 0;
    dp.rgvarg = NULL;
    pDocument->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                      &dp, NULL, NULL, NULL);
  }
  // 退出Word
  propertyName = _wcsdup(L"Quit");
  hr = pWordApp->GetIDsOfNames(IID_NULL, &propertyName, 1, LOCALE_USER_DEFAULT,
                               &dispid);
  if (SUCCEEDED(hr)) {
    dp.cArgs = 0;
    dp.rgvarg = NULL;
    pWordApp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD,
                     &dp, NULL, NULL, NULL);
  }

  pDocument->Release();
  pDocuments->Release();
  pWordApp->Release();
  auto end = std::chrono::high_resolution_clock::now();
  auto duration =
      std::chrono::duration_cast<std::chrono::microseconds>(end - start);

  std::cout << "函数执行时间: " << duration.count() << " 微秒" << std::endl;
  return 0;
}