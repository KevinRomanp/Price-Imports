codeunit 50500 "PreciosExcelColumns"
{
    procedure GetTipoVenta(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Text

    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then
            exit(Buffer."Cell Value as Text");
    end;


    procedure GetCodDivisa(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        cv: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(cv, Buffer."Cell Value as Text");
            exit(cv);
        end;
    end;

    procedure GetAssignTo(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        cv: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(cv, Buffer."Cell Value as Text");
            exit(cv);
        end;
    end;

    procedure GetDate(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Date
    var
        d: Date;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(D, Buffer."Cell Value as Text");
            exit(D);
        end;
    end;


    procedure GetNoProducto(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        np: Code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(np, Buffer."Cell Value as Text");
            exit(np);
        end;
    end;

    procedure GetCodUnidadMedida(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        um: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(um, Buffer."Cell Value as Text");
            exit(um);
        end;
    end;

    procedure GetCantidadMinima(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Decimal
    var
        cm: Decimal;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(cm, Buffer."Cell Value as Text");
            exit(cm);
        end;

    end;

    procedure GetAmountType(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        um: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(um, Buffer."Cell Value as Text");
            exit(um);
        end;
    end;

    procedure GetPrecioUnitario(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Decimal
    var
        pu: Decimal;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(pu, Buffer."Cell Value as Text");
            exit(pu);
        end;

    end;

    procedure GetPorcientoLinea(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Decimal
    var
        pu: Decimal;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(pu, Buffer."Cell Value as Text");
            exit(pu);
        end;

    end;

    procedure GetImporteDescuento(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Decimal
    var
        pu: Decimal;
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(pu, Buffer."Cell Value as Text");
            exit(pu);
        end;
    end;

    procedure GetShippingAddress(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        um: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(um, Buffer."Cell Value as Text");
            exit(um);
        end;
    end;

    procedure GetVatPostingGroup(var Buffer: Record "Excel Buffer" temporary; Col: Text; Row: Integer): Code[20]
    var
        cv: code[20];
    begin
        if Buffer.Get(Row, GetColumnNumber(col)) then begin
            Evaluate(cv, Buffer."Cell Value as Text");
            exit(cv);
        end;
    end;

    procedure GetColumnNumber(ColumnName: Text): Integer
    var
        columnIndex: Integer;
        factor: Integer;
        pos: Integer;
    begin
        factor := 1;
        for pos := strlen(ColumnName) downto 1 do
            if ColumnName[pos] >= 65 then begin
                columnIndex += factor * ((ColumnName[pos] - 65) + 1);
                factor *= 26;
            end;

        exit(columnIndex);
    end;

    procedure GetColumnName(columnNumber: Integer): Text
    var
        dividend: Integer;
        columnName: Text;
        modulo: Integer;
        c: Char;
    begin
        dividend := columnNumber;

        while (dividend > 0) do begin
            modulo := (dividend - 1) mod 26;
            c := 65 + modulo;
            columnName := format(c) + columnName;
            dividend := (dividend - modulo) DIV 26;
        end;

        exit(columnName);
    end;

}
