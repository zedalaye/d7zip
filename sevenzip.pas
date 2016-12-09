(********************************************************************************)
(*                        7-ZIP DELPHI API                                      *)
(*                                                                              *)
(* The contents of this file are subject to the Mozilla Public License Version  *)
(* 1.1 (the "License"); you may not use this file except in compliance with the *)
(* License. You may obtain a copy of the License at http://www.mozilla.org/MPL/ *)
(*                                                                              *)
(* Software distributed under the License is distributed on an "AS IS" basis,   *)
(* WITHOUT WARRANTY OF ANY KIND, either express or implied. See the License for *)
(* the specific language governing rights and limitations under the License.    *)
(*                                                                              *)
(* Unit owner : Henri Gourvest <hgourvest@gmail.com>                            *)
(* V1.2                                                                         *)
(********************************************************************************)

unit sevenzip;
{$ALIGN ON}
{$MINENUMSIZE 4}
{$WARN SYMBOL_PLATFORM OFF}	

interface
uses SysUtils, Windows, ActiveX, Classes, Contnrs;

type
  PVarType = ^TVarType;
  PCardArray = ^TCardArray;
  TCardArray = array[0..MaxInt div SizeOf(Cardinal) - 1] of Cardinal;

{$IFNDEF UNICODE}
  UnicodeString = WideString;
{$ENDIF}

//******************************************************************************
// PropID.h
//******************************************************************************

const
  kpidNoProperty            = 0;
  kpidMainSubfile           = 1;
  kpidHandlerItemIndex      = 2;
  kpidPath                  = 3;  // VT_BSTR
  kpidName                  = 4;  // VT_BSTR
  kpidExtension             = 5;  // VT_BSTR
  kpidIsDir                 = 6;  // VT_BOOL
  kpidSize                  = 7;  // VT_UI8
  kpidPackSize              = 8;  // VT_UI8
  kpidAttrib                = 9;  // VT_UI4
  kpidCTime                 = 10; // VT_FILETIME
  kpidATime                 = 11; // VT_FILETIME
  kpidMTime                 = 12; // VT_FILETIME
  kpidSolid                 = 13; // VT_BOOL
  kpidCommented             = 14; // VT_BOOL
  kpidEncrypted             = 15; // VT_BOOL
  kpidSplitBefore           = 16; // VT_BOOL
  kpidSplitAfter            = 17; // VT_BOOL
  kpidDictionarySize        = 18; // VT_UI4
  kpidCRC                   = 19; // VT_UI4
  kpidType                  = 20; // VT_BSTR
  kpidIsAnti                = 21; // VT_BOOL
  kpidMethod                = 22; // VT_BSTR
  kpidHostOS                = 23; // VT_BSTR
  kpidFileSystem            = 24; // VT_BSTR
  kpidUser                  = 25; // VT_BSTR
  kpidGroup                 = 26; // VT_BSTR
  kpidBlock                 = 27; // VT_UI4
  kpidComment               = 28; // VT_BSTR
  kpidPosition              = 29; // VT_UI4
  kpidPrefix                = 30; // VT_BSTR
  kpidNumSubDirs            = 31; // VT_UI4
  kpidNumSubFiles           = 32; // VT_UI4
  kpidUnpackVer             = 33; // VT_UI1
  kpidVolume                = 34; // VT_UI4
  kpidIsVolume              = 35; // VT_BOOL
  kpidOffset                = 36; // VT_UI8
  kpidLinks                 = 37; // VT_UI4
  kpidNumBlocks             = 38; // VT_UI4
  kpidNumVolumes            = 39; // VT_UI4
  kpidTimeType              = 40; // VT_UI4
  kpidBit64                 = 41; // VT_BOOL
  kpidBigEndian             = 42; // VT_BOOL
  kpidCpu                   = 43; // VT_BSTR
  kpidPhySize               = 44; // VT_UI8
  kpidHeadersSize           = 45; // VT_UI8
  kpidChecksum              = 46; // VT_UI4
  kpidCharacts              = 47; // VT_BSTR
  kpidVa                    = 48; // VT_UI8
  kpidId                    = 49;
  kpidShortName             = 50;
  kpidCreatorApp            = 51;
  kpidSectorSize            = 52;
  kpidPosixAttrib           = 53;
  kpidSymLink               = 54;
  kpidError                 = 55;
  kpidTotalSize             = 56;
  kpidFreeSpace             = 57;
  kpidClusterSize           = 58;
  kpidVolumeName            = 59;
  kpidLocalName             = 60;
  kpidProvider              = 61;
  kpidNtSecure              = 62;
  kpidIsAltStream           = 63;
  kpidIsAux                 = 64;
  kpidIsDeleted             = 65;
  kpidIsTree                = 66;
  kpidSha1                  = 67;
  kpidSha256                = 68;
  kpidErrorType             = 69;
  kpidNumErrors             = 70;
  kpidErrorFlags            = 71;
  kpidWarningFlags          = 72;
  kpidWarning               = 73;
  kpidNumStreams            = 74;
  kpidNumAltStreams         = 75;
  kpidAltStreamsSize        = 76;
  kpidVirtualSize           = 77;
  kpidUnpackSize            = 78;
  kpidTotalPhySize          = 79;
  kpidVolumeIndex           = 80;
  kpidSubType               = 81;
  kpidShortComment          = 82;
  kpidCodePage              = 83;
  kpidIsNotArcType          = 84;
  kpidPhySizeCantBeDetected = 85;
  kpidZerosTailIsAllowed    = 86;
  kpidTailSize              = 87;
  kpidEmbeddedStubSize      = 88;
  kpidNtReparse             = 89;
  kpidHardLink              = 90;
  kpidINode                 = 91;
  kpidStreamId              = 92;
  kpidReadOnly              = 93;
  kpidOutName               = 94;
  kpidCopyLink              = 95;

//  kpidTotalSize        = $1100; // VT_UI8
//  kpidFreeSpace        = kpidTotalSize + 1; // VT_UI8
//  kpidClusterSize      = kpidFreeSpace + 1; // VT_UI8
//  kpidVolumeName       = kpidClusterSize + 1; // VT_BSTR
//
//  kpidLocalName        = $1200; // VT_BSTR
//  kpidProvider         = kpidLocalName + 1; // VT_BSTR

  kpidUserDefined      = $10000;

//******************************************************************************
// IProgress.h
//******************************************************************************
type
  IProgress = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000000050000}']
    function SetTotal(total: Int64): HRESULT; stdcall;
    function SetCompleted(completeValue: PInt64): HRESULT; stdcall;
  end;

//******************************************************************************
// IPassword.h
//******************************************************************************

  ICryptoGetTextPassword = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000500100000}']
    function CryptoGetTextPassword(var password: TBStr): HRESULT; stdcall;
  end;

  ICryptoGetTextPassword2 = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000500110000}']
    function CryptoGetTextPassword2(passwordIsDefined: PInteger; var password: TBStr): HRESULT; stdcall;
  end;

//******************************************************************************
// IStream.h
// "23170F69-40C1-278A-0000-000300xx0000"
//******************************************************************************

  ISequentialInStream = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300010000}']
    function Read(data: Pointer; size: Cardinal; processedSize: PCardinal): HRESULT; stdcall;
    (*
    The requirement for caller: (processedSize != NULL).
    The callee can allow (processedSize == NULL) for compatibility reasons.

    if (size == 0), this function returns S_OK and (*processedSize) is set to 0.

    if (size != 0)
    {
      Partial read is allowed: (*processedSize <= avail_size && *processedSize <= size),
        where (avail_size) is the size of remaining bytes in stream.
      If (avail_size != 0), this function must read at least 1 byte: (*processedSize > 0).
      You must call Read() in loop, if you need to read exact amount of data.
    }

    If seek pointer before Read() call was changed to position past the end of stream:
      if (seek_pointer >= stream_size), this function returns S_OK and (*processedSize) is set to 0.

    ERROR CASES:
      If the function returns error code, then (*processedSize) is size of
      data written to (data) buffer (it can be data before error or data with errors).
      The recommended way for callee to work with reading errors:
        1) write part of data before error to (data) buffer and return S_OK.
        2) return error code for further calls of Read().
    *)
  end;

  ISequentialOutStream = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300020000}']
    function Write(data: Pointer; size: Cardinal; processedSize: PCardinal): HRESULT; stdcall;
    (*
    The requirement for caller: (processedSize != NULL).
    The callee can allow (processedSize == NULL) for compatibility reasons.

    if (size != 0)
    {
      Partial write is allowed: (*processedSize <= size),
      but this function must write at least 1 byte: (*processedSize > 0).
      You must call Write() in loop, if you need to write exact amount of data.
    }

    ERROR CASES:
      If the function returns error code, then (*processedSize) is size of
      data written from (data) buffer.
    *)
  end;

  IInStream = interface(ISequentialInStream)
  ['{23170F69-40C1-278A-0000-000300030000}']
    function Seek(offset: Int64; seekOrigin: Cardinal; newPosition: PInt64): HRESULT; stdcall;
  end;

  IOutStream = interface(ISequentialOutStream)
  ['{23170F69-40C1-278A-0000-000300040000}']
    function Seek(offset: Int64; seekOrigin: Cardinal; newPosition: PInt64): HRESULT; stdcall;
    function SetSize(newSize: Int64): HRESULT; stdcall;
  end;

  IStreamGetSize = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300060000}']
    function GetSize(size: PInt64): HRESULT; stdcall;
  end;

  IOutStreamFinish = interface(IUnknown)
  ['{23170F69-40C1-278A-0000-000300070000}']
    function Flush: HRESULT; stdcall;
  end;

//******************************************************************************
// IArchive.h
//******************************************************************************

// MIDL_INTERFACE("23170F69-40C1-278A-0000-000600xx0000")
//#define ARCHIVE_INTERFACE_SUB(i, base,  x) \
//DEFINE_GUID(IID_ ## i, \
//0x23170F69, 0x40C1, 0x278A, 0x00, 0x00, 0x00, 0x06, 0x00, x, 0x00, 0x00); \
//struct i: public base

//#define ARCHIVE_INTERFACE(i, x) ARCHIVE_INTERFACE_SUB(i, IUnknown, x)

type
// NFileTimeType
  NFileTimeType = (
    kWindows = 0,
    kUnix,
    kDOS
  );

// NArcInfoFlags
  NArcInfoFlags = (
    aifKeepName        = 1 shl 0,  // keep name of file in archive name
    aifAltStreams      = 1 shl 1,  // the handler supports alt streams
    aifNtSecure        = 1 shl 2,  // the handler supports NT security
    aifFindSignature   = 1 shl 3,  // the handler can find start of archive
    aifMultiSignature  = 1 shl 4,  // there are several signatures
    aifUseGlobalOffset = 1 shl 5,  // the seek position of stream must be set as global offset
    aifStartOpen       = 1 shl 6,  // call handler for each start position
    aifPureStartOpen   = 1 shl 7,  // call handler only for start of file
    aifBackwardOpen    = 1 shl 8,  // archive can be open backward
    aifPreArc          = 1 shl 9,  // such archive can be stored before real archive (like SFX stub)
    aifSymLinks        = 1 shl 10, // the handler supports symbolic links
    aifHardLinks       = 1 shl 11  // the handler supports hard links
  );

// NArchive::NHandlerPropID
  NHandlerPropID = (
    kName = 0,          // VT_BSTR
    kClassID,           // binary GUID in VT_BSTR
    kExtension,         // VT_BSTR
    kAddExtension,      // VT_BSTR
    kUpdate,            // VT_BOOL
    kKeepName,          // VT_BOOL
    kSignature,         // binary in VT_BSTR
    kMultiSignature,    // binary in VT_BSTR
    kSignatureOffset,   // VT_UI4
    kAltStreams,        // VT_BOOL
    kNtSecure,          // VT_BOOL
    kFlags              // VT_UI4
    // kVersion           // VT_UI4 ((VER_MAJOR << 8) | VER_MINOR)
  );

// NArchive::NExtract::NAskMode
  NAskMode = (
    kExtract = 0,
    kTest,
    kSkip
  );

// NArchive::NExtract::NOperationResult
  NExtOperationResult = (
    kOK = 0,
    kUnSupportedMethod,
    kDataError,
    kCRCError,
    kUnavailable,
    kUnexpectedEnd,
    kDataAfterEnd,
    kIsNotArc,
    kHeadersError,
    kWrongPassword
  );

// NArchive::NEventIndexType
  NEventIndexType = (
    kNoIndex = 0,
    kInArcIndex,
    kBlockIndex,
    kOutArcIndex
  );

// NArchive::NUpdate::NOperationResult
  NUpdOperationResult = (
    kOK_   = 0,
    kError
  );

  IArchiveOpenCallback = interface
  ['{23170F69-40C1-278A-0000-000600100000}']
    function SetTotal(files, bytes: PInt64): HRESULT; stdcall;
    function SetCompleted(files, bytes: PInt64): HRESULT; stdcall;
    (*
    IArchiveExtractCallback::

    7-Zip doesn't call IArchiveExtractCallback functions
      GetStream()
      PrepareOperation()
      SetOperationResult()
    from different threads simultaneously.
    But 7-Zip can call functions for IProgress or ICompressProgressInfo functions
    from another threads simultaneously with calls for IArchiveExtractCallback interface.

    IArchiveExtractCallback::GetStream()
      UInt32 index - index of item in Archive
      Int32 askExtractMode  (Extract::NAskMode)
        if (askMode != NExtract::NAskMode::kExtract)
        {
          then the callee can not real stream: (*inStream == NULL)
        }

      Out:
          (*inStream == NULL) - for directories
          (*inStream == NULL) - if link (hard link or symbolic link) was created
          if (*inStream == NULL && askMode == NExtract::NAskMode::kExtract)
          {
            then the caller must skip extracting of that file.
          }

      returns:
        S_OK     : OK
        S_FALSE  : data error (for decoders)

    if (IProgress::SetTotal() was called)
    {
      IProgress::SetCompleted(completeValue) uses
        packSize   - for some stream formats (xz, gz, bz2, lzma, z, ppmd).
        unpackSize - for another formats.
    }
    else
    {
      IProgress::SetCompleted(completeValue) uses packSize.
    }

    SetOperationResult()
      7-Zip calls SetOperationResult at the end of extracting,
      so the callee can close the file, set attributes, timestamps and security information.

      Int32 opRes (NExtract::NOperationResult)
    *)
  end;

  IArchiveExtractCallback = interface(IProgress)
  ['{23170F69-40C1-278A-0000-000600200000}']
    function GetStream(index: Cardinal; var outStream: ISequentialOutStream;
        askExtractMode: NAskMode): HRESULT; stdcall;
    // GetStream OUT: S_OK - OK, S_FALSE - skeep this file
    function PrepareOperation(askExtractMode: NAskMode): HRESULT; stdcall;
    function SetOperationResult(resultEOperationResult: NExtOperationResult): HRESULT; stdcall;
    (*
    IArchiveExtractCallbackMessage can be requested from IArchiveExtractCallback object
      by Extract() or UpdateItems() functions to report about extracting errors
    ReportExtractResult()
      UInt32 indexType (NEventIndexType)
      UInt32 index
      Int32 opRes (NExtract::NOperationResult)
    *)
  end;

  IArchiveExtractCallbackMessage = interface
  ['{23170F69-40C1-278A-0000-000600210000}']
    function ReportExtractResult(indexType: NEventIndexType; index: Cardinal; opRes: Integer): HRESULT; stdcall;
  end;

  IArchiveOpenVolumeCallback = interface
  ['{23170F69-40C1-278A-0000-000600300000}']
    function GetProperty(propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function GetStream(const name: PWideChar; var inStream: IInStream): HRESULT; stdcall;
  end;

  IInArchiveGetStream = interface
  ['{23170F69-40C1-278A-0000-000600400000}']
    function GetStream(index: Cardinal; var stream: ISequentialInStream ): HRESULT; stdcall;
  end;

  IArchiveOpenSetSubArchiveName = interface
  ['{23170F69-40C1-278A-0000-000600500000}']
    function SetSubArchiveName(name: PWideChar): HRESULT; stdcall;
  end;

  IInArchive = interface
  ['{23170F69-40C1-278A-0000-000600600000}']
    function Open(stream: IInStream; const maxCheckStartPosition: PInt64;
        openArchiveCallback: IArchiveOpenCallback): HRESULT; stdcall;
    (*
    IInArchive::Open
        stream
          if (kUseGlobalOffset), stream current position can be non 0.
          if (!kUseGlobalOffset), stream current position is 0.
        if (maxCheckStartPosition == NULL), the handler can try to search archive start in stream
        if (*maxCheckStartPosition == 0), the handler must check only current position as archive start

    IInArchive::Extract:
      indices must be sorted
      numItems = (UInt32)(Int32)-1 = 0xFFFFFFFF means "all files"
      testMode != 0 means "test files without writing to outStream"

    IInArchive::GetArchiveProperty:
      kpidOffset  - start offset of archive.
          VT_EMPTY : means offset = 0.
          VT_UI4, VT_UI8, VT_I8 : result offset; negative values is allowed
      kpidPhySize - size of archive. VT_EMPTY means unknown size.
        kpidPhySize is allowed to be larger than file size. In that case it must show
        supposed size.

      kpidIsDeleted:
      kpidIsAltStream:
      kpidIsAux:
      kpidINode:
        must return VARIANT_TRUE (VT_BOOL), if archive can support that property in GetProperty.


    Notes:
      Don't call IInArchive functions for same IInArchive object from different threads simultaneously.
      Some IInArchive handlers will work incorrectly in that case.
    *)

    function Close: HRESULT; stdcall;
    function GetNumberOfItems(var numItems: CArdinal): HRESULT; stdcall;
    function GetProperty(index: Cardinal; propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function Extract(indices: PCardArray; numItems: Cardinal;
        testMode: Integer; extractCallback: IArchiveExtractCallback): HRESULT; stdcall;
    // indices must be sorted
    // numItems = 0xFFFFFFFF means all files
    // testMode != 0 means "test files operation"

    function GetArchiveProperty(propID: PROPID; var value: OleVariant): HRESULT; stdcall;

    function GetNumberOfProperties(numProperties: PCardinal): HRESULT; stdcall;
    function GetPropertyInfo(index: Cardinal;
        name: PBSTR; propID: PPropID; varType: PVarType): HRESULT; stdcall;

    function GetNumberOfArchiveProperties(var numProperties: Cardinal): HRESULT; stdcall;
    function GetArchivePropertyInfo(index: Cardinal;
        name: PBSTR; propID: PPropID; varType: PVARTYPE): HRESULT; stdcall;
  end;

  IArchiveOpenSeq = interface
  ['{23170F69-40C1-278A-0000-000600610000}']
    function OpenSeq(var stream: ISequentialInStream): HRESULT; stdcall;
  end;

// NParentType::
  NParentType = (
    kDir = 0,
    kAltStream
  );

// NPropDataType::
  NPropDataType = (
    kMask_ZeroEnd   = $10, // 1 shl 4,
    // kMask_BigEndian = 1 shl 5,
    kMask_Utf       = $40, // 1 shl 6,
    kMask_Utf8      = $40, // kMask_Utf or 0,
    kMask_Utf16     = $41, // kMask_Utf or 1,
    // kMask_Utf32 = $42, // kMask_Utf or 2,

    kNotDefined = 0,
    kRaw = 1,

    kUtf8z  = $50, // kMask_Utf8  or kMask_ZeroEnd,
    kUtf16z = $51  // kMask_Utf16 or kMask_ZeroEnd
  );

  IArchiveGetRawProps = interface
  ['{23170F69-40C1-278A-0000-000600700000}']
    function GetParent(index: Cardinal; var parent: Cardinal; var parentType: Cardinal): HRESULT; stdcall;
    function GetRawProp(index: Cardinal; propID: PROPID; var data: Pointer; var dataSize: Cardinal; var propType: Cardinal): HRESULT; stdcall;
    function GetNumRawProps(var numProps: Cardinal): HRESULT; stdcall;
    function GetRawPropInfo(index: Cardinal; name: PBSTR; var propID: PROPID): HRESULT; stdcall;
  end;

  IArchiveGetRootProps = interface
  ['{23170F69-40C1-278A-0000-000600710000}']
    function GetRootProp(propID: PROPID; var value: PROPVARIANT): HRESULT; stdcall;
    function GetRootRawProp(propID: PROPID; var data: Pointer; var dataSize: Cardinal; var propType: Cardinal): HRESULT; stdcall;
  end;

  IArchiveUpdateCallback = interface(IProgress)
  ['{23170F69-40C1-278A-0000-000600800000}']
    function GetUpdateItemInfo(index: Cardinal;
        newData: PInteger; // 1 - new data, 0 - old data
        newProperties: PInteger; // 1 - new properties, 0 - old properties
        indexInArchive: PCardinal // -1 if there is no in archive, or if doesn't matter
        ): HRESULT; stdcall;
    function GetProperty(index: Cardinal; propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function GetStream(index: Cardinal; var inStream: ISequentialInStream): HRESULT; stdcall;
    function SetOperationResult(operationResult: Integer): HRESULT; stdcall;
  end;

  IArchiveUpdateCallback2 = interface(IArchiveUpdateCallback)
  ['{23170F69-40C1-278A-0000-000600820000}']
    function GetVolumeSize(index: Cardinal; size: PInt64): HRESULT; stdcall;
    function GetVolumeStream(index: Cardinal; var volumeStream: ISequentialOutStream): HRESULT; stdcall;
  end;

  IOutArchive = interface
  ['{23170F69-40C1-278A-0000-000600A00000}']
    function UpdateItems(outStream: ISequentialOutStream; numItems: Cardinal;
      updateCallback: IArchiveUpdateCallback): HRESULT; stdcall;
    function GetFileTimeType(type_: PCardinal): HRESULT; stdcall;
  end;

  ISetProperties = interface
  ['{23170F69-40C1-278A-0000-000600030000}']
    function SetProperties(names: PPWideChar; values: PPROPVARIANT; numProperties: Integer): HRESULT; stdcall;
  end;

//******************************************************************************
// ICoder.h
// "23170F69-40C1-278A-0000-000400xx0000"
//******************************************************************************

  ICompressProgressInfo = interface
  ['{23170F69-40C1-278A-0000-000400040000}']
    function SetRatioInfo(inSize, outSize: PInt64): HRESULT; stdcall;
  end;

  ICompressCoder = interface
  ['{23170F69-40C1-278A-0000-000400050000}']
  function Code(inStream, outStream: ISequentialInStream;
      inSize, outSize: PInt64;
      progress: ICompressProgressInfo): HRESULT; stdcall;
  end;

  ICompressCoder2 = interface
  ['{23170F69-40C1-278A-0000-000400180000}']
  function Code(var inStreams: ISequentialInStream;
      var inSizes: PInt64;
      numInStreams: Cardinal;
      var outStreams: ISequentialOutStream;
      var outSizes: PInt64;
      numOutStreams: Cardinal;
      progress: ICompressProgressInfo): HRESULT; stdcall;
  end;

//NCoderPropID::
  NCoderPropID = (
    kDefaultProp = 0,
    kDictionarySize,
    kUsedMemorySize,
    kOrder,
    kBlockSize,
    kPosStateBits,
    kLitContextBits,
    kLitPosBits,
    kNumFastBytes,
    kMatchFinder,
    kMatchFinderCycles,
    kNumPasses,
    kAlgorithm,
    kNumThreads,
    kEndMarker,
    kLevel,
    kReduceSize // estimated size of data that will be compressed. Encoder can use this value to reduce dictionary size.
  );

type
  ICompressSetCoderProperties = interface
  ['{23170F69-40C1-278A-0000-000400200000}']
    function SetCoderProperties(propIDs: PPropID;
      properties: PROPVARIANT; numProperties: Cardinal): HRESULT; stdcall;
  end;

(*
CODER_INTERFACE(ICompressSetCoderProperties, 0x21)
{
  STDMETHOD(SetDecoderProperties)(ISequentialInStream *inStream) PURE;
};
*)

  ICompressSetDecoderProperties2 = interface
  ['{23170F69-40C1-278A-0000-000400220000}']
    function SetDecoderProperties2(data: PByte; size: Cardinal): HRESULT; stdcall;
  end;

  ICompressWriteCoderProperties = interface
  ['{23170F69-40C1-278A-0000-000400230000}']
    function WriteCoderProperties(outStreams: ISequentialOutStream): HRESULT; stdcall;
  end;

  ICompressGetInStreamProcessedSize = interface
  ['{23170F69-40C1-278A-0000-000400240000}']
    function GetInStreamProcessedSize(value: PInt64): HRESULT; stdcall;
  end;

  ICompressSetCoderMt = interface
  ['{23170F69-40C1-278A-0000-000400250000}']
    function SetNumberOfThreads(numThreads: Cardinal): HRESULT; stdcall;
  end;

  ICompressGetSubStreamSize = interface
  ['{23170F69-40C1-278A-0000-000400300000}']
    function GetSubStreamSize(subStream: Int64; value: PInt64): HRESULT; stdcall;
  end;

  ICompressSetInStream = interface
  ['{23170F69-40C1-278A-0000-000400310000}']
    function SetInStream(inStream: ISequentialInStream): HRESULT; stdcall;
    function ReleaseInStream: HRESULT; stdcall;
  end;

  ICompressSetOutStream = interface
  ['{23170F69-40C1-278A-0000-000400320000}']
    function SetOutStream(outStream: ISequentialOutStream): HRESULT; stdcall;
    function ReleaseOutStream: HRESULT; stdcall;
  end;

  ICompressSetInStreamSize = interface
  ['{23170F69-40C1-278A-0000-000400330000}']
    function SetInStreamSize(inSize: PInt64): HRESULT; stdcall;
  end;

  ICompressSetOutStreamSize = interface
  ['{23170F69-40C1-278A-0000-000400340000}']
    function SetOutStreamSize(outSize: PInt64): HRESULT; stdcall;
  end;

  ICompressFilter = interface
  ['{23170F69-40C1-278A-0000-000400400000}']
    function Init: HRESULT; stdcall;
    function Filter(data: PByte; size: Cardinal): Cardinal; stdcall;
    // Filter return outSize (Cardinal)
    // if (outSize <= size): Filter have converted outSize bytes
    // if (outSize > size): Filter have not converted anything.
    //      and it needs at least outSize bytes to convert one block
    //      (it's for crypto block algorithms).
  end;

  ICryptoProperties = interface
  ['{23170F69-40C1-278A-0000-000400800000}']
    function SetKey(Data: PByte; size: Cardinal): HRESULT; stdcall;
    function SetInitVector(data: PByte; size: Cardinal): HRESULT; stdcall;
  end;

  ICryptoSetPassword = interface
  ['{23170F69-40C1-278A-0000-000400900000}']
    function CryptoSetPassword(data: PByte; size: Cardinal): HRESULT; stdcall;
  end;

  ICryptoSetCRC = interface
  ['{23170F69-40C1-278A-0000-000400A00000}']
    function CryptoSetCRC(crc: Cardinal): HRESULT; stdcall;
  end;

//////////////////////
// It's for DLL file
//NMethodPropID::
  NMethodPropID = (
    kID,
    kMethodName, // kName
    kDecoder,
    kEncoder,
    kPackStreams,
    kUnpackStreams,
    kDescription,
    kDecoderIsAssigned,
    kEncoderIsAssigned,
    kDigestSize
  );

//******************************************************************************
// CLASSES
//******************************************************************************

  T7zPasswordCallback = function(sender: Pointer; var password: UnicodeString): HRESULT; stdcall;
  T7zGetStreamCallBack = function(sender: Pointer; index: Cardinal;
    var outStream: ISequentialOutStream): HRESULT; stdcall;
  T7zProgressCallback = function(sender: Pointer; total: boolean; value: int64): HRESULT; stdcall;

  I7zInArchive = interface
  ['{022CF785-3ECE-46EF-9755-291FA84CC6C9}']
    procedure OpenFile(const filename: string); stdcall;
    procedure OpenStream(stream: IInStream); stdcall;
    procedure Close; stdcall;
    function GetNumberOfItems: Cardinal; stdcall;
    function GetItemPath(const index: integer): UnicodeString; stdcall;
    function GetItemName(const index: integer): UnicodeString; stdcall;
    function GetItemSize(const index: integer): Cardinal; stdcall;
    function GetItemIsFolder(const index: integer): boolean; stdcall;
    function GetInArchive: IInArchive;
    procedure ExtractItem(const item: Cardinal; Stream: TStream; test: longbool); stdcall;
    procedure ExtractItems(items: PCardArray; count: cardinal; test: longbool;
      sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure ExtractAll(test: longbool; sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure ExtractTo(const path: string); stdcall;
    procedure SetPasswordCallback(sender: Pointer; callback: T7zPasswordCallback); stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure SetClassId(const classid: TGUID);
    function GetClassId: TGUID;
    property ClassId: TGUID read GetClassId write SetClassId;
    property NumberOfItems: Cardinal read GetNumberOfItems;
    property ItemPath[const index: integer]: UnicodeString read GetItemPath;
    property ItemName[const index: integer]: UnicodeString read GetItemName;
    property ItemSize[const index: integer]: Cardinal read GetItemSize;
    property ItemIsFolder[const index: integer]: boolean read GetItemIsFolder;
    property InArchive: IInArchive read GetInArchive;
  end;

  I7zOutArchive = interface
  ['{BAA9D5DC-9FF4-4382-9BFD-EC9065BD0125}']
    procedure AddStream(Stream: TStream; Ownership: TStreamOwnership; Attributes: Cardinal;
      CreationTime, LastWriteTime: TFileTime; const Path: UnicodeString;
      IsFolder, IsAnti: boolean); stdcall;
    procedure AddFile(const Filename: TFileName; const Path: UnicodeString); stdcall;
    procedure AddFiles(const Dir, Path, Wildcard: string; recurse: boolean); stdcall;
    procedure SaveToFile(const FileName: TFileName); stdcall;
    procedure SaveToStream(stream: TStream); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure ClearBatch; stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    procedure SetPropertie(name: UnicodeString; value: OleVariant); stdcall;
    procedure SetClassId(const classid: TGUID);
    function GetClassId: TGUID;
    property ClassId: TGUID read GetClassId write SetClassId;
  end;

  I7zCodec = interface
  ['{AB48F772-F6B1-411E-907F-1567DB0E93B3}']

  end;


  T7zStream = class(TInterfacedObject, IInStream, IStreamGetSize,
    ISequentialOutStream, ISequentialInStream, IOutStream, IOutStreamFinish)
  private
    FStream: TStream;
    FOwnership: TStreamOwnership;
  protected
    function Read(data: Pointer; size: Cardinal; processedSize: PCardinal): HRESULT; stdcall;
    function Seek(offset: Int64; seekOrigin: Cardinal; newPosition: Pint64): HRESULT; stdcall;
    function GetSize(size: PInt64): HRESULT; stdcall;
    function SetSize(newSize: Int64): HRESULT; stdcall;
    function Write(data: Pointer; size: Cardinal; processedSize: PCardinal): HRESULT; stdcall;
    function Flush: HRESULT; stdcall;
  public
    constructor Create(Stream: TStream; Ownership: TStreamOwnership = soReference);
    destructor Destroy; override;
  end;

  // I7zOutArchive property setters
type
  TZipCompressionMethod = (mzCopy, mzDeflate, mzDeflate64, mzBZip2, mzLZMA, mzPPMD);
  TZipEncryptionMethod = (emAES128, emAES192, emAES256, emZIPCRYPTO);
  T7zCompressionMethod = (m7Copy, m7LZMA, m7BZip2, m7PPMd, m7Deflate, m7Deflate64);
                                                                                              //  ZIP 7z GZIP BZ2
  procedure SetCompressionLevel(Arch: I7zOutArchive; level: Cardinal);                        //   X   X   X   X
  procedure SetMultiThreading(Arch: I7zOutArchive; ThreadCount: Cardinal);                    //   X   X       X

  procedure SetCompressionMethod(Arch: I7zOutArchive; method: TZipCompressionMethod);         //   X
  procedure SetEncryptionMethod(Arch: I7zOutArchive; method: TZipEncryptionMethod);           //   X
  procedure SetDictionnarySize(Arch: I7zOutArchive; size: Cardinal); // < 32                  //   X           X
  procedure SetMemorySize(Arch: I7zOutArchive; size: Cardinal);                               //   X
  procedure SetDeflateNumPasses(Arch: I7zOutArchive; pass: Cardinal);                         //   X       X   X
  procedure SetNumFastBytes(Arch: I7zOutArchive; fb: Cardinal);                               //   X       X
  procedure SetNumMatchFinderCycles(Arch: I7zOutArchive; mc: Cardinal);                       //   X       X


  procedure SevenZipSetCompressionMethod(Arch: I7zOutArchive; method: T7zCompressionMethod);  //       X
  procedure SevenZipSetBindInfo(Arch: I7zOutArchive; const bind: UnicodeString);              //       X
  procedure SevenZipSetSolidSettings(Arch: I7zOutArchive; solid: boolean);                    //       X
  procedure SevenZipRemoveSfxBlock(Arch: I7zOutArchive; remove: boolean);                     //       X
  procedure SevenZipAutoFilter(Arch: I7zOutArchive; auto: boolean);                           //       X
  procedure SevenZipCompressHeaders(Arch: I7zOutArchive; compress: boolean);                  //       X
  procedure SevenZipCompressHeadersFull(Arch: I7zOutArchive; compress: boolean);              //       X
  procedure SevenZipEncryptHeaders(Arch: I7zOutArchive; Encrypt: boolean);                    //       X
  procedure SevenZipVolumeMode(Arch: I7zOutArchive; Mode: boolean);                           //       X

  // filetime util functions
  function DateTimeToFileTime(dt: TDateTime): TFileTime;
  function FileTimeToDateTime(ft: TFileTime): TDateTime;
  function CurrentFileTime: TFileTime;

  // constructors

  function CreateInArchive(const classid: TGUID; const lib: string = '7z.dll'): I7zInArchive;
  function CreateOutArchive(const classid: TGUID; const lib: string = '7z.dll'): I7zOutArchive;

const
  CLSID_CFormatZip      : TGUID = '{23170F69-40C1-278A-1000-000110010000}'; // [OUT] zip jar xpi
  CLSID_CFormatBZ2      : TGUID = '{23170F69-40C1-278A-1000-000110020000}'; // [OUT] bz2 bzip2 tbz2 tbz
  CLSID_CFormatRar      : TGUID = '{23170F69-40C1-278A-1000-000110030000}'; // [IN ] rar r00
  CLSID_CFormatArj      : TGUID = '{23170F69-40C1-278A-1000-000110040000}'; // [IN ] arj
  CLSID_CFormatZ        : TGUID = '{23170F69-40C1-278A-1000-000110050000}'; // [IN ] z taz
  CLSID_CFormatLzh      : TGUID = '{23170F69-40C1-278A-1000-000110060000}'; // [IN ] lzh lha
  CLSID_CFormat7z       : TGUID = '{23170F69-40C1-278A-1000-000110070000}'; // [OUT] 7z
  CLSID_CFormatCab      : TGUID = '{23170F69-40C1-278A-1000-000110080000}'; // [IN ] cab
  CLSID_CFormatNsis     : TGUID = '{23170F69-40C1-278A-1000-000110090000}'; // [IN ] nsis
  CLSID_CFormatLzma     : TGUID = '{23170F69-40C1-278A-1000-0001100A0000}'; // [IN ] lzma
  CLSID_CFormatLzma86   : TGUID = '{23170F69-40C1-278A-1000-0001100B0000}'; // [IN ] lzma 86
  CLSID_CFormatXz       : TGUID = '{23170F69-40C1-278A-1000-0001100C0000}'; // [OUT] xz
  CLSID_CFormatPpmd     : TGUID = '{23170F69-40C1-278A-1000-0001100D0000}'; // [IN ] ppmd

  CLSID_CFormatExt      : TGUID = '{23170F69-40C1-278A-1000-000110C70000}'; // [IN ] ext
  CLSID_CFormatVMDK     : TGUID = '{23170F69-40C1-278A-1000-000110C80000}'; // [IN ] vmdk
  CLSID_CFormatVDI      : TGUID = '{23170F69-40C1-278A-1000-000110C90000}'; // [IN ] vdi
  CLSID_CFormatQcow     : TGUID = '{23170F69-40C1-278A-1000-000110CA0000}'; // [IN ] qcow
  CLSID_CFormatGPT      : TGUID = '{23170F69-40C1-278A-1000-000110CB0000}'; // [IN ] GPT
  CLSID_CFormatRar5     : TGUID = '{23170F69-40C1-278A-1000-000110CC0000}'; // [IN ] Rar5
  CLSID_CFormatIHex     : TGUID = '{23170F69-40C1-278A-1000-000110CD0000}'; // [IN ] IHex
  CLSID_CFormatHxs      : TGUID = '{23170F69-40C1-278A-1000-000110CE0000}'; // [IN ] Hxs
  CLSID_CFormatTE       : TGUID = '{23170F69-40C1-278A-1000-000110CF0000}'; // [IN ] TE
  CLSID_CFormatUEFIc    : TGUID = '{23170F69-40C1-278A-1000-000110D00000}'; // [IN ] UEFIc
  CLSID_CFormatUEFIs    : TGUID = '{23170F69-40C1-278A-1000-000110D10000}'; // [IN ] UEFIs
  CLSID_CFormatSquashFS : TGUID = '{23170F69-40C1-278A-1000-000110D20000}'; // [IN ] SquashFS
  CLSID_CFormatCramFS   : TGUID = '{23170F69-40C1-278A-1000-000110D30000}'; // [IN ] CramFS
  CLSID_CFormatAPM      : TGUID = '{23170F69-40C1-278A-1000-000110D40000}'; // [IN ] APM
  CLSID_CFormatMslz     : TGUID = '{23170F69-40C1-278A-1000-000110D50000}'; // [IN ] MsLZ
  CLSID_CFormatFlv      : TGUID = '{23170F69-40C1-278A-1000-000110D60000}'; // [IN ] FLV
  CLSID_CFormatSwf      : TGUID = '{23170F69-40C1-278A-1000-000110D70000}'; // [IN ] SWF
  CLSID_CFormatSwfc     : TGUID = '{23170F69-40C1-278A-1000-000110D80000}'; // [IN ] SWFC
  CLSID_CFormatNtfs     : TGUID = '{23170F69-40C1-278A-1000-000110D90000}'; // [IN ] NTFS
  CLSID_CFormatFat      : TGUID = '{23170F69-40C1-278A-1000-000110DA0000}'; // [IN ] FAT
  CLSID_CFormatMbr      : TGUID = '{23170F69-40C1-278A-1000-000110DB0000}'; // [IN ] MBR
  CLSID_CFormatVhd      : TGUID = '{23170F69-40C1-278A-1000-000110DC0000}'; // [IN ] VHD
  CLSID_CFormatPe       : TGUID = '{23170F69-40C1-278A-1000-000110DD0000}'; // [IN ] PE (Windows Exe)
  CLSID_CFormatElf      : TGUID = '{23170F69-40C1-278A-1000-000110DE0000}'; // [IN ] ELF (Linux Exe)
  CLSID_CFormatMachO    : TGUID = '{23170F69-40C1-278A-1000-000110DF0000}'; // [IN ] Mach-O
  CLSID_CFormatUdf      : TGUID = '{23170F69-40C1-278A-1000-000110E00000}'; // [IN ] iso
  CLSID_CFormatXar      : TGUID = '{23170F69-40C1-278A-1000-000110E10000}'; // [IN ] xar
  CLSID_CFormatMub      : TGUID = '{23170F69-40C1-278A-1000-000110E20000}'; // [IN ] mub
  CLSID_CFormatHfs      : TGUID = '{23170F69-40C1-278A-1000-000110E30000}'; // [IN ] HFS
  CLSID_CFormatDmg      : TGUID = '{23170F69-40C1-278A-1000-000110E40000}'; // [IN ] dmg
  CLSID_CFormatCompound : TGUID = '{23170F69-40C1-278A-1000-000110E50000}'; // [IN ] msi doc xls ppt
  CLSID_CFormatWim      : TGUID = '{23170F69-40C1-278A-1000-000110E60000}'; // [OUT] wim swm
  CLSID_CFormatIso      : TGUID = '{23170F69-40C1-278A-1000-000110E70000}'; // [IN ] iso
  CLSID_CFormatBkf      : TGUID = '{23170F69-40C1-278A-1000-000110E80000}'; // [IN ] BKF
  CLSID_CFormatChm      : TGUID = '{23170F69-40C1-278A-1000-000110E90000}'; // [IN ] chm chi chq chw hxs hxi hxr hxq hxw lit
  CLSID_CFormatSplit    : TGUID = '{23170F69-40C1-278A-1000-000110EA0000}'; // [IN ] 001
  CLSID_CFormatRpm      : TGUID = '{23170F69-40C1-278A-1000-000110EB0000}'; // [IN ] rpm
  CLSID_CFormatDeb      : TGUID = '{23170F69-40C1-278A-1000-000110EC0000}'; // [IN ] deb
  CLSID_CFormatCpio     : TGUID = '{23170F69-40C1-278A-1000-000110ED0000}'; // [IN ] cpio
  CLSID_CFormatTar      : TGUID = '{23170F69-40C1-278A-1000-000110EE0000}'; // [OUT] tar
  CLSID_CFormatGZip     : TGUID = '{23170F69-40C1-278A-1000-000110EF0000}'; // [OUT] gz gzip tgz tpz

implementation

const
  MAXCHECK : int64 = (1 shl 20);
  ZipCompressionMethod: array[TZipCompressionMethod] of UnicodeString = ('COPY', 'DEFLATE', 'DEFLATE64', 'BZIP2', 'LZMA', 'PPMD');
  ZipEncryptionMethod: array[TZipEncryptionMethod] of UnicodeString = ('AES128', 'AES192', 'AES256', 'ZIPCRYPTO');
  SevCompressionMethod: array[T7zCompressionMethod] of UnicodeString = ('COPY', 'LZMA', 'BZIP2', 'PPMD', 'DEFLATE', 'DEFLATE64');

function DateTimeToFileTime(dt: TDateTime): TFileTime;
var
  st: TSystemTime;
begin
  DateTimeToSystemTime(dt, st);
  if not (SystemTimeToFileTime(st, Result) and LocalFileTimeToFileTime(Result, Result))
    then RaiseLastOSError;
end;

function FileTimeToDateTime(ft: TFileTime): TDateTime;
var
  st: TSystemTime;
begin
  if not (FileTimeToLocalFileTime(ft, ft) and FileTimeToSystemTime(ft, st)) then
    RaiseLastOSError;
  Result := SystemTimeToDateTime(st);
end;

function CurrentFileTime: TFileTime;
begin
  GetSystemTimeAsFileTime(Result);
end;

procedure RINOK(const hr: HRESULT);
begin
  if hr <> S_OK then
    raise Exception.Create(SysErrorMessage(hr));
end;

procedure SetCardinalProperty(arch: I7zOutArchive; const name: UnicodeString; card: Cardinal);
var
  value: OleVariant;
begin
  TPropVariant(value).vt := VT_UI4;
  TPropVariant(value).ulVal := card;
  arch.SetPropertie(name, value);
end;

procedure SetBooleanProperty(arch: I7zOutArchive; const name: UnicodeString; bool: boolean);
begin
  case bool of
    true: arch.SetPropertie(name, 'ON');
    false: arch.SetPropertie(name, 'OFF');
  end;
end;

procedure SetCompressionLevel(Arch: I7zOutArchive; level: Cardinal);
begin
  SetCardinalProperty(arch, 'X', level);
end;

procedure SetMultiThreading(Arch: I7zOutArchive; ThreadCount: Cardinal);
begin
  SetCardinalProperty(arch, 'MT', ThreadCount);
end;

procedure SetCompressionMethod(Arch: I7zOutArchive; method: TZipCompressionMethod);
begin
  Arch.SetPropertie('M', ZipCompressionMethod[method]);
end;

procedure SetEncryptionMethod(Arch: I7zOutArchive; method: TZipEncryptionMethod);
begin
  Arch.SetPropertie('EM', ZipEncryptionMethod[method]);
end;

procedure SetDictionnarySize(Arch: I7zOutArchive; size: Cardinal);
begin
  SetCardinalProperty(arch, 'D', size);
end;

procedure SetMemorySize(Arch: I7zOutArchive; size: Cardinal);
begin
  SetCardinalProperty(arch, 'MEM', size);
end;

procedure SetDeflateNumPasses(Arch: I7zOutArchive; pass: Cardinal);
begin
  SetCardinalProperty(arch, 'PASS', pass);
end;

procedure SetNumFastBytes(Arch: I7zOutArchive; fb: Cardinal);
begin
  SetCardinalProperty(arch, 'FB', fb);
end;

procedure SetNumMatchFinderCycles(Arch: I7zOutArchive; mc: Cardinal);
begin
  SetCardinalProperty(arch, 'MC', mc);
end;

procedure SevenZipSetCompressionMethod(Arch: I7zOutArchive; method: T7zCompressionMethod);
begin
  Arch.SetPropertie('0', SevCompressionMethod[method]);
end;

procedure SevenZipSetBindInfo(Arch: I7zOutArchive; const bind: UnicodeString);
begin
  arch.SetPropertie('B', bind);
end;

procedure SevenZipSetSolidSettings(Arch: I7zOutArchive; solid: boolean);
begin
  SetBooleanProperty(Arch, 'S', solid);
end;

procedure SevenZipRemoveSfxBlock(Arch: I7zOutArchive; remove: boolean);
begin
  SetBooleanProperty(Arch, 'RSFX', remove);
end;

procedure SevenZipAutoFilter(Arch: I7zOutArchive; auto: boolean);
begin
  SetBooleanProperty(Arch, 'F', auto);
end;

procedure SevenZipCompressHeaders(Arch: I7zOutArchive; compress: boolean);
begin
  SetBooleanProperty(Arch, 'HC', compress);
end;

procedure SevenZipCompressHeadersFull(Arch: I7zOutArchive; compress: boolean);
begin
  SetBooleanProperty(arch, 'HCF', compress);
end;

procedure SevenZipEncryptHeaders(Arch: I7zOutArchive; Encrypt: boolean);
begin
  SetBooleanProperty(arch, 'HE', Encrypt);
end;

procedure SevenZipVolumeMode(Arch: I7zOutArchive; Mode: boolean);
begin
  SetBooleanProperty(arch, 'V', Mode);
end;

type
  T7zPlugin = class(TInterfacedObject)
  private
    FHandle: THandle;
    FCreateObject: function(const clsid, iid :TGUID; var outObject): HRESULT; stdcall;
  public
    constructor Create(const lib: string); virtual;
    destructor Destroy; override;
    procedure CreateObject(const clsid, iid :TGUID; var obj);
  end;

  T7zCodec = class(T7zPlugin, I7zCodec, ICompressProgressInfo)
  private
    FGetMethodProperty: function(index: Cardinal; propID: NMethodPropID; var value: OleVariant): HRESULT; stdcall;
    FGetNumberOfMethods: function(numMethods: PCardinal): HRESULT; stdcall;
    function GetNumberOfMethods: Cardinal;
    function GetMethodProperty(index: Cardinal; propID: NMethodPropID): OleVariant;
    function GetName(const index: integer): string;
  protected
    function SetRatioInfo(inSize, outSize: PInt64): HRESULT; stdcall;
  public
    function GetDecoder(const index: integer): ICompressCoder;
    function GetEncoder(const index: integer): ICompressCoder;
    constructor Create(const lib: string); override;
    property MethodProperty[index: Cardinal; propID: NMethodPropID]: OleVariant read GetMethodProperty;
    property NumberOfMethods: Cardinal read GetNumberOfMethods;
    property Name[const index: integer]: string read GetName;
  end;

  T7zArchive = class(T7zPlugin)
  private
    FGetHandlerProperty: function(propID: NHandlerPropID; var value: OleVariant): HRESULT; stdcall;
    FClassId: TGUID;
    procedure SetClassId(const classid: TGUID);
    function GetClassId: TGUID;
  public
    function GetHandlerProperty(const propID: NHandlerPropID): OleVariant;
    function GetLibStringProperty(const Index: NHandlerPropID): string;
    function GetLibGUIDProperty(const Index: NHandlerPropID): TGUID;
    constructor Create(const lib: string); override;
    property HandlerProperty[const propID: NHandlerPropID]: OleVariant read GetHandlerProperty;
    property Name: string index kName read GetLibStringProperty;
    property ClassID: TGUID read GetClassId write SetClassId;
    property Extension: string index kExtension read GetLibStringProperty;
  end;

  T7zInArchive = class(T7zArchive, I7zInArchive, IProgress, IArchiveOpenCallback,
    IArchiveExtractCallback, ICryptoGetTextPassword, IArchiveOpenVolumeCallback,
    IArchiveOpenSetSubArchiveName)
  private
    FInArchive: IInArchive;
    FPasswordCallback: T7zPasswordCallback;
    FPasswordSender: Pointer;
    FProgressCallback: T7zProgressCallback;
    FProgressSender: Pointer;
    FStream: TStream;
    FPasswordIsDefined: Boolean;
    FPassword: UnicodeString;
    FSubArchiveMode: Boolean;
    FSubArchiveName: UnicodeString;
    FExtractCallBack: T7zGetStreamCallBack;
    FExtractSender: Pointer;
    FExtractPath: string;
    function GetInArchive: IInArchive;
    function GetItemProp(const Item: Cardinal; prop: PROPID): OleVariant;
  protected
    // I7zInArchive
    procedure OpenFile(const filename: string); stdcall;
    procedure OpenStream(stream: IInStream); stdcall;
    procedure Close; stdcall;
    function GetNumberOfItems: Cardinal; stdcall;
    function GetItemPath(const index: integer): UnicodeString; stdcall;
    function GetItemName(const index: integer): UnicodeString; stdcall;
    function GetItemSize(const index: integer): Cardinal; stdcall; stdcall;
    function GetItemIsFolder(const index: integer): boolean; stdcall;
    procedure ExtractItem(const item: Cardinal; Stream: TStream; test: longbool); stdcall;
    procedure ExtractItems(items: PCardArray; count: cardinal; test: longbool; sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure SetPasswordCallback(sender: Pointer; callback: T7zPasswordCallback); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure ExtractAll(test: longbool; sender: pointer; callback: T7zGetStreamCallBack); stdcall;
    procedure ExtractTo(const path: string); stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    // IArchiveOpenCallback
    function SetTotal(files, bytes: PInt64): HRESULT; overload; stdcall;
    function SetCompleted(files, bytes: PInt64): HRESULT; overload; stdcall;
    // IProgress
    function SetTotal(total: Int64): HRESULT;  overload; stdcall;
    function SetCompleted(completeValue: PInt64): HRESULT; overload; stdcall;
    // IArchiveExtractCallback
    function GetStream(index: Cardinal; var outStream: ISequentialOutStream;
      askExtractMode: NAskMode): HRESULT; overload; stdcall;
    function PrepareOperation(askExtractMode: NAskMode): HRESULT; stdcall;
    function SetOperationResult(resultEOperationResult: NExtOperationResult): HRESULT; overload; stdcall;
    // ICryptoGetTextPassword
    function CryptoGetTextPassword(var password: TBStr): HRESULT; stdcall;
    // IArchiveOpenVolumeCallback
    function GetProperty(propID: PROPID; var value: OleVariant): HRESULT; overload; stdcall;
    function GetStream(const name: PWideChar; var inStream: IInStream): HRESULT; overload; stdcall;
    // IArchiveOpenSetSubArchiveName
    function SetSubArchiveName(name: PWideChar): HRESULT; stdcall;

  public
    constructor Create(const lib: string); override;
    destructor Destroy; override;
    property InArchive: IInArchive read GetInArchive;
  end;

  T7zOutArchive = class(T7zArchive, I7zOutArchive, IArchiveUpdateCallback, ICryptoGetTextPassword2)
  private
    FOutArchive: IOutArchive;
    FBatchList: TObjectList;
    FProgressCallback: T7zProgressCallback;
    FProgressSender: Pointer;
    FPassword: UnicodeString;
    function GetOutArchive: IOutArchive;
  protected
    // I7zOutArchive
    procedure AddStream(Stream: TStream; Ownership: TStreamOwnership;
      Attributes: Cardinal; CreationTime, LastWriteTime: TFileTime;
      const Path: UnicodeString; IsFolder, IsAnti: boolean); stdcall;
    procedure AddFile(const Filename: TFileName; const Path: UnicodeString); stdcall;
    procedure AddFiles(const Dir, Path, Wildcard: string; recurse: boolean); stdcall;
    procedure SaveToFile(const FileName: TFileName); stdcall;
    procedure SaveToStream(stream: TStream); stdcall;
    procedure SetProgressCallback(sender: Pointer; callback: T7zProgressCallback); stdcall;
    procedure ClearBatch; stdcall;
    procedure SetPassword(const password: UnicodeString); stdcall;
    procedure SetPropertie(name: UnicodeString; value: OleVariant); stdcall;
    // IProgress
    function SetTotal(total: Int64): HRESULT; stdcall;
    function SetCompleted(completeValue: PInt64): HRESULT; stdcall;
    // IArchiveUpdateCallback
    function GetUpdateItemInfo(index: Cardinal;
        newData: PInteger; // 1 - new data, 0 - old data
        newProperties: PInteger; // 1 - new properties, 0 - old properties
        indexInArchive: PCardinal // -1 if there is no in archive, or if doesn't matter
        ): HRESULT; stdcall;
    function GetProperty(index: Cardinal; propID: PROPID; var value: OleVariant): HRESULT; stdcall;
    function GetStream(index: Cardinal; var inStream: ISequentialInStream): HRESULT; stdcall;
    function SetOperationResult(operationResult: Integer): HRESULT; stdcall;
    // ICryptoGetTextPassword2
    function CryptoGetTextPassword2(passwordIsDefined: PInteger; var password: TBStr): HRESULT; stdcall;
  public
    constructor Create(const lib: string); override;
    destructor Destroy; override;
    property OutArchive: IOutArchive read GetOutArchive;
  end;

function CreateInArchive(const classid: TGUID; const lib: string): I7zInArchive;
begin
  Result := T7zInArchive.Create(lib);
  Result.ClassId := classid;
end;

function CreateOutArchive(const classid: TGUID; const lib: string): I7zOutArchive;
begin
  Result := T7zOutArchive.Create(lib);
  Result.ClassId := classid;
end;


{ T7zPlugin }

constructor T7zPlugin.Create(const lib: string);
begin
  FHandle := LoadLibrary(PChar(lib));
  if FHandle = 0 then
    raise exception.CreateFmt('Error loading library %s', [lib]);
  FCreateObject := GetProcAddress(FHandle, 'CreateObject');
  if not (Assigned(FCreateObject)) then
  begin
    FreeLibrary(FHandle);
    raise Exception.CreateFmt('%s is not a 7z library', [lib]);
  end;
end;

destructor T7zPlugin.Destroy;
begin
  FreeLibrary(FHandle);
  inherited;
end;

procedure T7zPlugin.CreateObject(const clsid, iid: TGUID; var obj);
var
  hr: HRESULT;
begin
  hr := FCreateObject(clsid, iid, obj);
  if failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;

{ T7zCodec }

constructor T7zCodec.Create(const lib: string);
begin
  inherited;
  FGetMethodProperty := GetProcAddress(FHandle, 'GetMethodProperty');
  FGetNumberOfMethods := GetProcAddress(FHandle, 'GetNumberOfMethods');
  if not (Assigned(FGetMethodProperty) and Assigned(FGetNumberOfMethods)) then
  begin
    FreeLibrary(FHandle);
    raise Exception.CreateFmt('%s is not a codec library', [lib]);
  end;
end;

function T7zCodec.GetDecoder(const index: integer): ICompressCoder;
var
  v: OleVariant;
begin
  v := MethodProperty[index, kDecoder];
  CreateObject(TPropVariant(v).puuid^, ICompressCoder, Result);
end;

function T7zCodec.GetEncoder(const index: integer): ICompressCoder;
var
  v: OleVariant;
begin
  v := MethodProperty[index, kEncoder];
  CreateObject(TPropVariant(v).puuid^, ICompressCoder, Result);
end;

function T7zCodec.GetMethodProperty(index: Cardinal;
  propID: NMethodPropID): OleVariant;
var
  hr: HRESULT;
begin
  hr := FGetMethodProperty(index, propID, Result);
  if Failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;

function T7zCodec.GetName(const index: integer): string;
begin
  Result := MethodProperty[index, kMethodName];
end;

function T7zCodec.GetNumberOfMethods: Cardinal;
var
  hr: HRESULT;
begin
  hr := FGetNumberOfMethods(@Result);
  if Failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;


function T7zCodec.SetRatioInfo(inSize, outSize: PInt64): HRESULT;
begin
  Result := S_OK;
end;

{ T7zInArchive }

procedure T7zInArchive.Close; stdcall;
begin
  FPasswordIsDefined := false;
  FSubArchiveMode := false;
  FInArchive.Close;
  FInArchive := nil;
end;

constructor T7zInArchive.Create(const lib: string);
begin
  inherited;
  FPasswordCallback := nil;
  FPasswordSender := nil;
  FPasswordIsDefined := false;
  FSubArchiveMode := false;
  FExtractCallBack := nil;
  FExtractSender := nil;
end;

destructor T7zInArchive.Destroy;
begin
  FInArchive := nil;
  inherited;
end;

function T7zInArchive.GetInArchive: IInArchive;
begin
  if FInArchive = nil then
    CreateObject(ClassID, IInArchive, FInArchive);
  Result := FInArchive;
end;

function T7zInArchive.GetItemPath(const index: integer): UnicodeString; stdcall;
begin
  Result := UnicodeString(GetItemProp(index, kpidPath));
end;

function T7zInArchive.GetNumberOfItems: Cardinal; stdcall;
begin
  RINOK(FInArchive.GetNumberOfItems(Result));
end;

procedure T7zInArchive.OpenFile(const filename: string); stdcall;
var
  strm: IInStream;
begin
  strm := T7zStream.Create(TFileStream.Create(filename, fmOpenRead or fmShareDenyNone), soOwned);
  try
    RINOK(
      InArchive.Open(
        strm,
          @MAXCHECK, self as IArchiveOpenCallBack
        )
      );
  finally
    strm := nil;
  end;
end;

procedure T7zInArchive.OpenStream(stream: IInStream); stdcall;
begin
  RINOK(InArchive.Open(stream, @MAXCHECK, self as IArchiveOpenCallBack));
end;

function T7zInArchive.GetItemIsFolder(const index: integer): boolean; stdcall;
begin
  Result := Boolean(GetItemProp(index, kpidIsDir));
end;

function T7zInArchive.GetItemProp(const Item: Cardinal;
  prop: PROPID): OleVariant;
begin
  FInArchive.GetProperty(Item, prop, Result);
end;

procedure T7zInArchive.ExtractItem(const item: Cardinal; Stream: TStream; test: longbool); stdcall;
begin
  FStream := Stream;
  try
    if test then
      RINOK(FInArchive.Extract(@item, 1, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(@item, 1, 0, self as IArchiveExtractCallback));
  finally
    FStream := nil;
  end;
end;

function T7zInArchive.GetStream(index: Cardinal;
  var outStream: ISequentialOutStream; askExtractMode: NAskMode): HRESULT;
var
  path: string;
begin
  if askExtractMode = kExtract then
    if FStream <> nil then
      outStream := T7zStream.Create(FStream, soReference) as ISequentialOutStream else
    if assigned(FExtractCallback) then
    begin
      Result := FExtractCallBack(FExtractSender, index, outStream);
      Exit;
    end else
    if FExtractPath <> '' then
    begin
      if not GetItemIsFolder(index) then
      begin
        path := FExtractPath + GetItemPath(index);
        ForceDirectories(ExtractFilePath(path));
        outStream := T7zStream.Create(TFileStream.Create(path, fmCreate), soOwned);
      end;
    end;
  Result := S_OK;
end;

function T7zInArchive.PrepareOperation(askExtractMode: NAskMode): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.SetCompleted(completeValue: PInt64): HRESULT;
begin
  if Assigned(FProgressCallback) and (completeValue <> nil) then
    Result := FProgressCallback(FProgressSender, false, completeValue^) else
    Result := S_OK;
end;

function T7zInArchive.SetCompleted(files, bytes: PInt64): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.SetOperationResult(
  resultEOperationResult: NExtOperationResult): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.SetTotal(total: Int64): HRESULT;
begin
  if Assigned(FProgressCallback) then
    Result := FProgressCallback(FProgressSender, true, total) else
    Result := S_OK;
end;

function T7zInArchive.SetTotal(files, bytes: PInt64): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.CryptoGetTextPassword(var password: TBStr): HRESULT;
var
  wpass: UnicodeString;
begin
  if FPasswordIsDefined then
  begin
    password := SysAllocString(PWideChar(FPassword));
    Result := S_OK;
  end else
  if Assigned(FPasswordCallback) then
  begin
    Result := FPasswordCallBack(FPasswordSender, wpass);
    if Result = S_OK then
    begin
      password := SysAllocString(PWideChar(wpass));
      FPasswordIsDefined := True;
      FPassword := wpass;
    end;
  end else
    Result := S_FALSE;
end;

function T7zInArchive.GetProperty(propID: PROPID;
  var value: OleVariant): HRESULT;
begin
  Result := S_OK;
end;

function T7zInArchive.GetStream(const name: PWideChar;
  var inStream: IInStream): HRESULT;
begin
  Result := S_OK;
end;

procedure T7zInArchive.SetPasswordCallback(sender: Pointer;
  callback: T7zPasswordCallback); stdcall;
begin
  FPasswordSender := sender;
  FPasswordCallback := callback;
end;

function T7zInArchive.SetSubArchiveName(name: PWideChar): HRESULT;
begin
  FSubArchiveMode := true;
  FSubArchiveName := name;
  Result := S_OK;
end;

function T7zInArchive.GetItemName(const index: integer): UnicodeString; stdcall;
begin
  Result := UnicodeString(GetItemProp(index, kpidName));
end;

function T7zInArchive.GetItemSize(const index: integer): Cardinal; stdcall;
begin
  Result := Cardinal(GetItemProp(index, kpidSize));
end;

procedure T7zInArchive.ExtractItems(items: PCardArray; count: cardinal; test: longbool;
  sender: pointer; callback: T7zGetStreamCallBack); stdcall;
begin
  FExtractCallBack := callback;
  FExtractSender := sender;
  try
    if test then
      RINOK(FInArchive.Extract(items, count, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(items, count, 0, self as IArchiveExtractCallback));
  finally
    FExtractCallBack := nil;
    FExtractSender := nil;
  end;
end;

procedure T7zInArchive.SetProgressCallback(sender: Pointer;
  callback: T7zProgressCallback); stdcall;
begin
  FProgressSender := sender;
  FProgressCallback := callback;
end;

procedure T7zInArchive.ExtractAll(test: longbool; sender: pointer;
  callback: T7zGetStreamCallBack);
begin
  FExtractCallBack := callback;
  FExtractSender := sender;
  try
    if test then
      RINOK(FInArchive.Extract(nil, $FFFFFFFF, 1, self as IArchiveExtractCallback)) else
      RINOK(FInArchive.Extract(nil, $FFFFFFFF, 0, self as IArchiveExtractCallback));
  finally
    FExtractCallBack := nil;
    FExtractSender := nil;
  end;
end;

procedure T7zInArchive.ExtractTo(const path: string);
begin
  FExtractPath := IncludeTrailingPathDelimiter(path);
  try
    RINOK(FInArchive.Extract(nil, $FFFFFFFF, 0, self as IArchiveExtractCallback));
  finally
    FExtractPath := '';
  end;
end;

procedure T7zInArchive.SetPassword(const password: UnicodeString);
begin
  FPassword := password;
  FPasswordIsDefined :=  FPassword <> '';
end;

{ T7zArchive }

constructor T7zArchive.Create(const lib: string);
begin
  inherited;
  FGetHandlerProperty := GetProcAddress(FHandle, 'GetHandlerProperty');
  if not Assigned(FGetHandlerProperty) then
  begin
    FreeLibrary(FHandle);
    raise Exception.CreateFmt('%s is not a Format library', [lib]);
  end;
  FClassId := GUID_NULL;
end;

function T7zArchive.GetClassId: TGUID;
begin
  Result := FClassId;
end;

function T7zArchive.GetHandlerProperty(const propID: NHandlerPropID): OleVariant;
var
  hr: HRESULT;
begin
  hr := FGetHandlerProperty(propID, Result);
  if Failed(hr) then
    raise Exception.Create(SysErrorMessage(hr));
end;

function T7zArchive.GetLibGUIDProperty(const Index: NHandlerPropID): TGUID;
var
  v: OleVariant;
begin
  v := HandlerProperty[index];
  Result := TPropVariant(v).puuid^;
end;

function T7zArchive.GetLibStringProperty(const Index: NHandlerPropID): string;
begin
  Result := HandlerProperty[Index];
end;

procedure T7zArchive.SetClassId(const classid: TGUID);
begin
  FClassId := classid;
end;

{ T7zStream }

constructor T7zStream.Create(Stream: TStream; Ownership: TStreamOwnership);
begin
  inherited Create;
  FStream := Stream;
  FOwnership := Ownership;
end;

destructor T7zStream.destroy;
begin
  if FOwnership = soOwned then
  begin
    FStream.Free;
    FStream := nil;
  end;
  inherited;
end;

function T7zStream.Flush: HRESULT;
begin
  Result := S_OK;
end;

function T7zStream.GetSize(size: PInt64): HRESULT;
begin
  if size <> nil then
    size^ := FStream.Size;
  Result := S_OK;
end;

function T7zStream.Read(data: Pointer; size: Cardinal;
  processedSize: PCardinal): HRESULT;
var
  len: integer;
begin
  len := FStream.Read(data^, size);
  if processedSize <> nil then
    processedSize^ := len;
  Result := S_OK;
end;

function T7zStream.Seek(offset: Int64; seekOrigin: Cardinal;
  newPosition: PInt64): HRESULT;
begin
  FStream.Seek(offset, TSeekOrigin(seekOrigin));
  if newPosition <> nil then
    newPosition^ := FStream.Position;
  Result := S_OK;
end;

function T7zStream.SetSize(newSize: Int64): HRESULT;
begin
  FStream.Size := newSize;
  Result := S_OK;
end;

function T7zStream.Write(data: Pointer; size: Cardinal;
  processedSize: PCardinal): HRESULT;
var
  len: integer;
begin
  len := FStream.Write(data^, size);
  if processedSize <> nil then
    processedSize^ := len;
  Result := S_OK;
end;

type
  TSourceMode = (smStream, smFile);

  T7zBatchItem = class
    SourceMode: TSourceMode;
    Stream: TStream;
    Attributes: Cardinal;
    CreationTime, LastWriteTime: TFileTime;
    Path: UnicodeString;
    IsFolder, IsAnti: boolean;
    FileName: TFileName;
    Ownership: TStreamOwnership;
    Size: Cardinal;
    destructor Destroy; override;
  end;

destructor T7zBatchItem.Destroy;
begin
  if (Ownership = soOwned) and (Stream <> nil) then
    Stream.Free;
  inherited;
end;

{ T7zOutArchive }

procedure T7zOutArchive.AddFile(const Filename: TFileName; const Path: UnicodeString);
var
  item: T7zBatchItem;
  Handle: THandle;
begin
  if not FileExists(Filename) then exit;
  item := T7zBatchItem.Create;
  Item.SourceMode := smFile;
  item.Stream := nil;
  item.FileName := Filename;
  item.Path := Path;
  Handle := FileOpen(Filename, fmOpenRead or fmShareDenyNone);
  GetFileTime(Handle, @item.CreationTime, nil, @item.LastWriteTime);
  item.Size := GetFileSize(Handle, nil);
  CloseHandle(Handle);
  item.Attributes := GetFileAttributes(PChar(Filename));
  item.IsFolder := false;
  item.IsAnti := False;
  item.Ownership := soOwned;
  FBatchList.Add(item);
end;

procedure T7zOutArchive.AddFiles(const Dir, Path, Wildcard: string; recurse: boolean);
var
  lencut: integer;
  willlist: TStringList;
  zedir: string;
  procedure Traverse(p: string);
  var
    f: TSearchRec;
    i: integer;
    item: T7zBatchItem;
  begin
    if recurse then
    begin
      if FindFirst(p + '*.*', faDirectory, f) = 0 then
      repeat
        if (f.Name[1] <> '.') then
          Traverse(IncludeTrailingPathDelimiter(p + f.Name));
      until FindNext(f) <> 0;
      SysUtils.FindClose(f);
    end;

    for i := 0 to willlist.Count - 1 do
    begin
      if FindFirst(p + willlist[i], faReadOnly or faHidden or faSysFile or faArchive, f) = 0 then
      repeat
        item := T7zBatchItem.Create;
        Item.SourceMode := smFile;
        item.Stream := nil;
        item.FileName := p + f.Name;
        item.Path := copy(item.FileName, lencut, length(item.FileName) - lencut + 1);
        if path <> '' then
          item.Path := IncludeTrailingPathDelimiter(path) + item.Path;
        item.CreationTime := f.FindData.ftCreationTime;
        item.LastWriteTime := f.FindData.ftLastWriteTime;
        item.Attributes := f.FindData.dwFileAttributes;
        item.Size := f.Size;
        item.IsFolder := false;
        item.IsAnti := False;
        item.Ownership := soOwned;
        FBatchList.Add(item);
      until FindNext(f) <> 0;
      SysUtils.FindClose(f);
    end;
  end;
begin
  willlist := TStringList.Create;
  try
    willlist.Delimiter := ';';
    willlist.DelimitedText := Wildcard;
    zedir := IncludeTrailingPathDelimiter(Dir);
    lencut := Length(zedir) + 1;
    Traverse(zedir);
  finally
    willlist.Free;
  end;
end;

procedure T7zOutArchive.AddStream(Stream: TStream; Ownership: TStreamOwnership;
  Attributes: Cardinal; CreationTime, LastWriteTime: TFileTime;
  const Path: UnicodeString; IsFolder, IsAnti: boolean); stdcall;
var
  item: T7zBatchItem;
begin
  item := T7zBatchItem.Create;
  Item.SourceMode := smStream;
  item.Attributes := Attributes;
  item.CreationTime := CreationTime;
  item.LastWriteTime := LastWriteTime;
  item.Path := Path;
  item.IsFolder := IsFolder;
  item.IsAnti := IsAnti;
  item.Stream := Stream;
  item.Size := Stream.Size;
  item.Ownership := Ownership;
  FBatchList.Add(item);
end;

procedure T7zOutArchive.ClearBatch;
begin
  FBatchList.Clear;
end;

constructor T7zOutArchive.Create(const lib: string);
begin
  inherited;
  FBatchList := TObjectList.Create;
  FProgressCallback := nil;
  FProgressSender := nil;
end;

function T7zOutArchive.CryptoGetTextPassword2(passwordIsDefined: PInteger;
  var password: TBStr): HRESULT;
begin
  if FPassword <> '' then
  begin
   passwordIsDefined^ := 1;
   password := SysAllocString(PWideChar(FPassword));
  end else
    passwordIsDefined^ := 0;
  Result := S_OK;
end;

destructor T7zOutArchive.Destroy;
begin
  FOutArchive := nil;
  FBatchList.Free;
  inherited;
end;

function T7zOutArchive.GetOutArchive: IOutArchive;
begin
  if FOutArchive = nil then
    CreateObject(ClassID, IOutArchive, FOutArchive);
  Result := FOutArchive;
end;

function T7zOutArchive.GetProperty(index: Cardinal; propID: PROPID;
  var value: OleVariant): HRESULT;
var
  item: T7zBatchItem;
begin
  item := T7zBatchItem(FBatchList[index]);
  case propID of
    kpidAttrib:
      begin
        TPropVariant(Value).vt := VT_UI4;
        TPropVariant(Value).ulVal := item.Attributes;
      end;
    kpidMTime:
      begin
        TPropVariant(value).vt := VT_FILETIME;
        TPropVariant(value).filetime := item.LastWriteTime;
      end;
    kpidPath:
      begin
        if item.Path <> '' then
          value := item.Path;
      end;
    kpidIsDir: Value := item.IsFolder;
    kpidSize:
      begin
        TPropVariant(Value).vt := VT_UI8;
        TPropVariant(Value).uhVal.QuadPart := item.Size;
      end;
    kpidCTime:
      begin
        TPropVariant(value).vt := VT_FILETIME;
        TPropVariant(value).filetime := item.CreationTime;
      end;
    kpidIsAnti: value := item.IsAnti;
  else
   // beep(0,0);
  end;
  Result := S_OK;
end;

function T7zOutArchive.GetStream(index: Cardinal;
  var inStream: ISequentialInStream): HRESULT;
var
  item: T7zBatchItem;
begin
  item := T7zBatchItem(FBatchList[index]);
  case item.SourceMode of
    smFile: inStream := T7zStream.Create(TFileStream.Create(item.FileName, fmOpenRead or fmShareDenyNone), soOwned);
    smStream:
      begin
        item.Stream.Seek(0, soFromBeginning);
        inStream := T7zStream.Create(item.Stream);
      end;
  end;
  Result := S_OK;
end;

function T7zOutArchive.GetUpdateItemInfo(index: Cardinal; newData,
  newProperties: PInteger; indexInArchive: PCardinal): HRESULT;
begin
  newData^ := 1;
  newProperties^ := 1;
  indexInArchive^ := CArdinal(-1);
  Result := S_OK;
end;

procedure T7zOutArchive.SaveToFile(const FileName: TFileName);
var
  f: TFileStream;
begin
  f := TFileStream.Create(FileName, fmCreate);
  try
    SaveToStream(f);
  finally
    f.free;
  end;
end;

procedure T7zOutArchive.SaveToStream(stream: TStream);
var
  strm: ISequentialOutStream;
begin
  strm := T7zStream.Create(stream);
  try
    RINOK(OutArchive.UpdateItems(strm, FBatchList.Count, self as IArchiveUpdateCallback));
  finally
    strm := nil;
  end;
end;

function T7zOutArchive.SetCompleted(completeValue: PInt64): HRESULT;
begin
  if Assigned(FProgressCallback) and (completeValue <> nil) then
    Result := FProgressCallback(FProgressSender, false, completeValue^) else
    Result := S_OK;
end;

function T7zOutArchive.SetOperationResult(
  operationResult: Integer): HRESULT;
begin
  Result := S_OK;
end;

procedure T7zOutArchive.SetPassword(const password: UnicodeString);
begin
  FPassword := password;
end;

procedure T7zOutArchive.SetProgressCallback(sender: Pointer;
  callback: T7zProgressCallback);
begin
  FProgressCallback := callback;
  FProgressSender := sender;
end;

procedure T7zOutArchive.SetPropertie(name: UnicodeString;
  value: OleVariant);
var
  intf: ISetProperties;
  p: PWideChar;
begin
  intf := OutArchive as ISetProperties;
  p := PWideChar(name);
  RINOK(intf.SetProperties(@p, @TPropVariant(value), 1));
end;

function T7zOutArchive.SetTotal(total: Int64): HRESULT;
begin
  if Assigned(FProgressCallback) then
    Result := FProgressCallback(FProgressSender, true, total) else
    Result := S_OK;
end;

end.
