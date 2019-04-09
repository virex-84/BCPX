{
  (с) https://github.com/virex-84
  
  Аналог утилиты bcp в MSSQL
  Позволяющий делать запрос к БД
  И сохранять результат в excel

  eebcpx.exe -Q"тело запроса" -S"SERVER-NAME" -F"C:\temp\text.xslx" -H -I -U"User" -P"Password"
  -Q - запрос
  -U - имя пользователя
  -P - пароль
  -F - имя файла
  -H - включать в файл заголовоки запроса/таблицы
  -S - имя сервера
  -I - интерактивный режим: каждую секунду отображает время и прогресс
}
program BCPX;

{$APPTYPE CONSOLE}

uses
  SysUtils,
  Classes,
  ADODB,
  ActiveX,
  Variants,
  Windows,
  ComObj;

//ключи программы
const
  pQuery = '-Q';
  pUser = '-U';
  pPassword = '-P';
  pFileName = '-F';
  pHeaders = '-H';
  pServer = '-S';
  pInteractive = '-I';

//импорт из Excel8TLB.pas
const
  xlInsideHorizontal = $0000000C;
  xlInsideVertical = $0000000B;
  xlEdgeBottom = $00000009;
  xlEdgeLeft = $00000007;
  xlEdgeRight = $0000000A;
  xlEdgeTop = $00000008;

  xlThin = $00000002;

//для определения типа колонки
const
  varDecimal = $000E;

var
  consoleLog: string = '';
  startTime: TDateTime;
  timer: WORD = 0;

  i: integer;
  param: string;

  query: string;
  user: string;
  password: string;
  filename: string;
  withHeaders: boolean;
  servername: string;
  isInteractive: boolean;

  connection: TADOConnection;
  recordset: _recordset;

  count: integer;
  Excel: Variant;

  ArrayData: variant;
  currentRange: variant;
  value: OleVariant;
  varExtended: Extended;

//таймер для интерактивного режима
function timeSetEvent(uDelay, uResolution: Longint; lpFunction: pointer; dwUser, uFlags: Longint): Longint; stdcall; external 'winmm.dll';
function timeKillEvent(uID: UINT): Integer; stdcall; external 'winmm';

//вытаскиваем текстовый параметр
procedure extractParam(var param: string; text: string; name: string); overload;
begin
  if copy(text, 1, length(name)) = name then
    param := copy(text, 1 + length(name), length(text));
end;

//вытаскиваем булевый параметр
procedure extractParam(var param: boolean; text: string; name: string); overload;
begin
  if copy(text, 1, length(name)) = name then
    param := true;
end;

//используется ли файл другой программой
function IsOpen(const aFileName: string): Boolean;
var
  Hf: Integer;
begin
  //Существует ли файл.
  Result := FileExists(aFileName);
  //Если файл не существует, значит он не открыт. Выходим.
  if not Result then Exit;
  //Проверяем, открыт ли уже файл. Для этого пытаемся открыть файл
  //в режиме неразделяемого доступа.
  Hf := FileOpen(aFileName, fmOpenReadWrite or fmShareExclusive);
  Result := Hf = -1;
  if not Result then FileClose(Hf);
end;

function StrToOem(AnsiStr: string): string;
begin
  SetLength(Result, Length(AnsiStr));
  if Length(Result) > 0 then
    CharToOemBuff(PChar(AnsiStr), PChar(Result), Length(Result));
end;

function OemToStr(const AnsiStr: string): string;
begin
  SetLength(Result, Length(AnsiStr));
  if Length(Result) > 0 then
    OemToAnsiBuff(PChar(AnsiStr), PChar(Result), Length(Result));
end;

//вывод лога в одной строке
procedure log(text: string; refreshNow: boolean = false);
var
  hStdOut: HWND;
  ScreenBufInfo: TConsoleScreenBufferInfo;
  Coord: TCoord;
  NumWritten: DWORD;

  time: string;
begin
  if not isInteractive then exit;

  consoleLog := text;

  if not refreshNow then exit;

  hStdOut := GetStdHandle(STD_OUTPUT_HANDLE);
  GetConsoleScreenBufferInfo(hStdOut, ScreenBufInfo);
  Coord.X := 0;
  Coord.Y := 0;

  //очищаем консоль
  FillConsoleOutputCharacter(hStdOut, ' ', ScreenBufInfo.dwSize.X * ScreenBufInfo.dwSize.Y, Coord, NumWritten);

  //установка позиции
  SetConsoleCursorPosition(hStdOut, Coord);

  time := FormatDateTime('hh:mm:ss', now() - startTime);

  //пишем текст
  Writeln(Format('Server: %s', [servername]));
  Writeln(Format('Query : %s', [StrToOem(query)]));
  Writeln(Format('Time  : %s', [time]));
  Writeln(Format('%s', [consoleLog]));
end;

//тик таймера
procedure OnTime(uTimerID, uMsg, dwUser, dw1, dw2: LongInt); stdcall;
begin
  log(consoleLog, true);
end;

//проверяем установлен ли на компьютере Excel
function isExcelInstalled: boolean;
var
  ClassID: TCLSID;
  Res: HRESULT;
begin
  Res := CLSIDFromProgID(PWideChar(WideString('Excel.Application')), ClassID);
  if Res = S_OK then
    Result := true
  else
    Result := false;
end;

begin
  //устанавливаем кодировку
  //SetConsoleCP(1251);
  //SetConsoleOutputCP(1251);

  //засекаем время запуска
  startTime := now();

  //извлекаем параметры
  for i := 0 to ParamCount do begin
    param := paramstr(i);

    extractParam(query, param, pQuery);
    extractParam(user, param, pUser);
    extractParam(password, param, pPassword);
    extractParam(filename, param, pFileName);
    extractParam(withHeaders, param, pHeaders);
    extractParam(servername, param, pServer);
    extractParam(isInteractive, param, pInteractive);
  end;

  //0 хендл окна
  //1 - хендл таймера
  //1 - интервал таймера
  //адресс процедуры @Proc
  if isInteractive then
    timer := timeSetEvent(1000, 1000, @OnTime, 0, 1);

  //не указали запрос, имя файла или имя сервера
  if (trim(query) = '') or (trim(filename) = '') or (servername = '') then exit;

  //проверяем доступен ли файл для записи
  //будет неудобно если сделали получасовой запрос, а файл для сохранения занят другой программой
  if IsOpen(filename) then begin
    Writeln(Format('Error open file "%s"', [filename]));
    exit;
  end;

  //если Excel не установлен
  if not isExcelInstalled then begin
    Writeln('Excel is not install');
    exit;
  end;

  try
    CoInitialize(nil);

    try
      Excel := CreateOleObject('Excel.Application');

      Excel.Visible := false;
      Excel.ScreenUpdating := false;
      Excel.EnableEvents := false;
      Excel.DisplayStatusBar := false;
      Excel.DisplayAlerts := false;
      //Excel.ActiveWindow.Caption := 'Title';

      //добавляем книгу
      Excel.WorkBooks.Add;

      //подключаемся к серверу базы данных
      connection := TADOConnection.Create(nil);
      connection.ConnectionString := Format('Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=RCReport;Data Source=%s;Current Language=Russian', [servername]);
      connection.Mode := cmRead;
      connection.CursorLocation := clUseServer;
      connection.IsolationLevel := ilChaos;
      connection.LoginPrompt := False;
      connection.CommandTimeout := 0; //ожидаем бесконечно (по умолчанию 30)

      //подключаемся к базе
      log('connection...', true);
      if ((user <> '') or (password <> '')) then
        connection.Open(user, password)
      else
        connection.Open;

      //начинаем транзакцию
      connection.BeginTrans;
      //получаем записи (ADO объект - "рекордсет") из БД, асинхронно
      log('execute...', true);
      recordset := connection.Execute(query, cmdText, [eoAsyncFetch]);

      //пробегаемся по записям
      count := 1;

      //создаем массив
      ArrayData := VarArrayCreate([0, 1 {recordset.RecordCount}, 0, recordset.Fields.Count], varVariant);

      //добавляем заголовки
      if withHeaders then begin
        //ArrayData := VarArrayCreate([0, 1{recordset.RecordCount}, 0, recordset.Fields.Count], varVariant);
        for i := 0 to recordset.Fields.Count - 1 do
          ArrayData[0, i] := VarToStr(recordset.Fields[i].Name);

        currentRange := Excel.Range[Excel.Cells.Item[1, 1], Excel.Cells.Item[VarArrayHighBound(ArrayData, 1) + 1, VarArrayHighBound(ArrayData, 2)]];
        currentRange.FormulaR1C1 := ArrayData;

        inc(count);
      end;


      //recordset.MoveFirst;
      while not (recordset.EOF) do begin
        log(Format('progress %d rows loaded', [count]));

        for i := 0 to recordset.Fields.Count - 1 do begin
          value := recordset.Fields[i].Value;

          //если это булево значение
          if VarIsType(value, [varBoolean]) then begin
            if value = true then
              value := 'Да'
            else
              value := 'Нет';
          end;

          //если это значение с плавающей запятой
          if VarIsType(value, [varSingle, varDouble, varCurrency, varDecimal]) then begin
            varExtended := value;
            value := varExtended;
          end;

          ArrayData[0, i] := value;

          VarClear(value);
        end;

        currentRange := Excel.Range[Excel.Cells.Item[count, 1], Excel.Cells.Item[VarArrayHighBound(ArrayData, 1) + count, VarArrayHighBound(ArrayData, 2)]];

        //вставка данных
        currentRange.FormulaR1C1 := ArrayData;

        inc(count);
        recordset.MoveNext;
      end;

      //освобождаем ресурсы
      //иначе процесс excel будет висеть в памяти
      VarClear(currentRange);
      VarClear(ArrayData);

      //выделяем всё
      currentRange := Excel.Range[Excel.Cells.Item[1, 1], Excel.Cells.Item[count - 1, recordset.Fields.Count]];

      //переносить по словам - очень медленно
      //currentRange.WrapText:=true;

      //рамка снаружи
      currentRange.Borders[xlEdgeBottom].Weight := xlThin;
      currentRange.Borders[xlEdgeLeft].Weight := xlThin;
      currentRange.Borders[xlEdgeRight].Weight := xlThin;
      currentRange.Borders[xlEdgeTop].Weight := xlThin;

      //рамка внутри
      currentRange.Borders[xlInsideHorizontal].Weight := xlThin;
      currentRange.Borders[xlInsideVertical].Weight := xlThin;

      //автофильтр для колонок
      if withHeaders then
        currentRange.AutoFilter;

      //авто-ширина и авто-высота
      Excel.ActiveWorkbook.Worksheets.Item[1].Columns.AutoFit;
      Excel.ActiveWorkbook.Worksheets.Item[1].Rows.AutoFit;

      //освобождаем ресурсы
      VarClear(currentRange);

      log(Format('progress %d rows loaded', [count]), true);

    except
      on e: Exception do Writeln(StrToOem(e.Message));
    end;

  finally
    //убиваем таймер
    if isInteractive then
      timeKillEvent(timer);

    //если рекордсет существует - закрываем
    (*
    if Assigned(recordset) then
      if (recordset.State>0{adStateClosed}) then begin
        recordset.Cancel;
        recordset.Close;
      end;
     *)

    //пытаемся сохранить
    if not VarIsNull(Excel) then
    try
      //сохраняем результат
      Excel.ActiveWorkbook.SaveAs(filename);
      //Excel.ActiveWorkbook.Close(true,'C:\\temp\\111.xlsx');

      //пишем в консоль результат
      //по этим строкам определяют результат работы bcp.exe от microsoft'а
      Writeln(Format('%d rows copied', [count]));

      //Excel.Visible := true;
      //Excel.ScreenUpdating := true;
      //Excel.EnableEvents := true;
      //Excel.DisplayStatusBar := true;
      Excel.Workbooks.Close;
      Excel.Quit;

      //освобождаем указатель на Excel
      VarClear(Excel);
    except
      on e: Exception do Writeln(StrToOem(e.Message));
    end;
  end;

  exit;

end.

