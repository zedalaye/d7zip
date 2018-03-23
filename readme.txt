官网已经找不到了。
这个地址比较新
https://github.com/zedalaye/d7zip
在这个基础上 融合了
SevenZip.pas BUG修改版 - 20160613 - 堕落恶魔 - 博客园
https://www.cnblogs.com/hs-kill/p/3876160.html
然后再加了一些小的修改。

最后，提供一个比较全面的 例子。
【Delphi】从内存读取或解压压缩文件（RAR、ZIP、TAR、GZIP等）（一） - 峋山隐修会 - 博客园
http://www.cnblogs.com/caibirdy1985/archive/2013/05/13/4232949.html
【Delphi】从内存读取或解压压缩文件（RAR、ZIP、TAR、GZIP等）（二） - 峋山隐修会 - 博客园
http://www.cnblogs.com/caibirdy1985/archive/2013/05/14/4232948.html



7-zip Delphi API
This API use the 7-zip dll (7z.dll) to read and write all 7-zip supported archive formats.

- Autor: Henri Gourvest <hgourvest@progdigy.com>
- Licence: MPL1.1
- Date: 15/04/2009
- Version: 1.1

Reading archive:

Extract to path:
解压到目录：
with CreateInArchive('Formats\zip.dll') do
begin
   OpenFile('c:\test.zip');
   ExtractTo('c:\test');
end;

Get file list:
获取文件列表:
with CreateInArchive('Formats\7z.dll') do
begin
   OpenFile('c:\test.7z');
   for i := 0 to NumberOfItems - 1 do
   if not ItemIsFolder[i] then
      Writeln(ItemPath[i]);
end;

Extract to stream
解压到流:
with CreateInArchive('Formats\7z.dll') do
begin
   OpenFile('c:\test.7z');
   for i := 0 to NumberOfItems - 1 do
    if not ItemIsFolder[i] then
       ExtractItem(i, stream, false);
end;

Extract "n" Items
解压多项目:
function GetStreamCallBack(sender: Pointer; index: Cardinal;
var outStream: ISequentialOutStream): HRESULT; stdcall;
begin
case index of ...
  outStream := T7zStream.Create(aStream, soReference);
  Result := S_OK;
end;

procedure TMainForm.ExtractClick(Sender: TObject);
var
i: integer;
items: array[0..2] of Cardinal;
begin
  with CreateInArchive('Formats\7z.dll') do
  begin
    OpenFile('c:\test.7z');
    // items must be sorted by index!
    items[0] := 0;
    items[1] := 1;
    items[2] := 2;
    ExtractItems(@items, Length(items), false, nil, GetStreamCallBack);
  end;
end;

Open stream
打开流:
with CreateInArchive('Formats\zip.dll') do
begin
   OpenStream(T7zStream.Create(TFileStream.Create('c:\test.zip', fmOpenRead), soOwned));
   OpenStream(aStream, soReference);
   ...
end;

Progress bar
进度条回调:
function ProgressCallback(sender: Pointer; total: boolean; value: int64): HRESULT; stdcall;
begin
   if total then
     Mainform.ProgressBar.Max := value else
     Mainform.ProgressBar.Position := value;
   Result := S_OK;
end;

procedure TMainForm.ExtractClick(Sender: TObject);
begin
   with CreateInArchive('Formats\zip.dll') do
begin
     OpenFile('c:\test.zip');
     SetProgressCallback(nil, ProgressCallback);
     ...
   end;
end;

Password
打开含有密码的文件:
function PasswordCallback(sender: Pointer; var password: WideString): HRESULT; stdcall;
begin
   // call a dialog box ...
   password := 'password';
   Result := S_OK;
end;
procedure TMainForm.ExtractClick(Sender: TObject);
begin
   with CreateInArchive('Formats\zip.dll') do
   begin
     // using callback
     SetPasswordCallback(nil, PasswordCallback);
     // or setting password directly
     SetPassword('password');
     OpenFile('c:\test.zip');
     ...
   end;
end;

Writing archive
压缩存档:
procedure TMainForm.ExtractAllClick(Sender: TObject);
var
   Arch: I7zOutArchive;
begin
   Arch := CreateOutArchive('Formats\7z.dll');
   // add a file
   Arch.AddFile('c:\test.bin', 'folder\test.bin');
   // add files using willcards and recursive search
   Arch.AddFiles('c:\test', 'folder', '*.pas;*.dfm', true, true);
   // add a stream
   Arch.AddStream(aStream, soReference, faArchive, CurrentFileTime, CurrentFileTime, 'folder\test.bin', false, false);
   // compression level
   SetCompressionLevel(Arch, 5);
   // compression method if <> LZMA
   SevenZipSetCompressionMethod(Arch, m7BZip2);
   // add a progress bar ...
   Arch.SetProgressCallback(...);
   // set a password if necessary
   Arch.SetPassword('password');
   // Save to file
   Arch.SaveToFile('c:\test.zip');
   // or a stream
   Arch.SaveToStream(aStream);
end;
