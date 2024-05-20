pageextension 50501 "Sales Price " extends "Sales Price List"
{
    actions
    {
        addlast(Processing)
        {
            action(ImportPrice)
            {
                Caption = 'Importar precio combustible';
                Image = ImportExcel;
                ApplicationArea = All;
                trigger OnAction()
                var
                    x: Codeunit PreciosExcelColumns;

                    PriceListLine: record "Price List Line";
                    PriceListLine2: record "Price List Line";
                    PriceListHeader: Record "Price List Header";
                    InS: InStream;
                    Filename: Text;
                    Row: Integer;
                    LastRow: Integer;
                    SheetName: Text[250];
                    SourceTypeText: Text;
                    SourceTypeEnum: enum "Price Source Type";
                    AmountTypeText: Text;
                    AmountTypeEnum: enum "Price Amount Type";
                    UploadErr: Label 'Please check if there are any empty values on the fields: Item No, Unit Measure Code or Starting Date.';
                    ErrorOrtografia: Label 'Por favor revise la ortografia en la columna "Tipo de Asignar a" en la línea %1 del Excel';
                    ErrorFecha: Label 'La fecha de vigencia no puede ser antes que %1. Revise la línea %2.';
                    ErrorFaltaParametro: Label 'El valor de la columna "%1" no puede estar en blanco en línea %2 del Excel.';
                    fecha: date;
                    TipoAsignarTxt: text[35];
                    NoAsignarTxt: text[35];
                    CodDivisaTxt: text[35];
                    FIText: text[35];
                    NoProductoTxt: text[35];

                    DefineTxt: text[35];
                    PrecioUnitarioTxt: text[35];

                    PorcientoDscLinTxt: text[35];
                    ImpDescTxt: text[35];
                    CodUnidadMedidaTxt: text[35];

                    TipoProductoTxt: text[35];
                    DireccionEnvioTxt: text[35];

                begin
                    Clear(TipoAsignarTxt);
                    Clear(NoAsignarTxt);
                    Clear(CodDivisaTxt);
                    Clear(FIText);
                    Clear(NoProductoTxt);

                    Clear(DefineTxt);
                    Clear(PrecioUnitarioTxt);

                    Clear(PorcientoDscLinTxt);
                    Clear(ImpDescTxt);
                    Clear(CodUnidadMedidaTxt);

                    Clear(SourceTypeText);
                    Clear(TipoProductoTxt);
                    Clear(DireccionEnvioTxt);

                    Buffer.DeleteAll();
                    if UploadIntoStream('Escoja un archivo', '', '', Filename, InS) then begin
                        if SheetName = '' then
                            SheetName := Buffer.SelectSheetsNameStream(InS);
                        Buffer.OpenBookStream(InS, SheetName);
                        Buffer.ReadSheet();
                        Buffer.setrange("Column No.", 1);
                        Buffer.FindLast();
                        LastRow := Buffer."Row No.";

                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Tipo de Asignar a');
                        if Buffer.FindFirst() then
                            TipoAsignarTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'N.º de Asignar a');
                        if Buffer.FindFirst() then
                            NoAsignarTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Cód. divisa');
                        if Buffer.FindFirst() then
                            CodDivisaTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Fecha inicial');
                        Buffer.FindFirst();
                        FIText := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Tipo de producto');
                        if Buffer.FindFirst() then
                            TipoProductoTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'N.º de producto');
                        Buffer.FindFirst();
                        NoProductoTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Cód. unidad medida');
                        if Buffer.FindFirst() then
                            CodUnidadMedidaTxt := buffer.xlColID;

                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Define');
                        if Buffer.FindFirst() then
                            DefineTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Precio unitario');
                        if Buffer.FindFirst() then
                            PrecioUnitarioTxt := buffer.xlColID;

                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", '% descuento línea');
                        if Buffer.FindFirst() then
                            PorcientoDscLinTxt := buffer.xlColID;
                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Importe Descuento');
                        if Buffer.FindFirst() then
                            ImpDescTxt := buffer.xlColID;

                        Buffer.Reset();
                        buffer.SetRange("Cell Value as Text", 'Dirección Envio');
                        if Buffer.FindFirst() then
                            DireccionEnvioTxt := buffer.xlColID;

                        buffer.Reset();
                        for row := 2 to LastRow do begin
                            SourceTypeText := x.GetTipoVenta(Buffer, TipoAsignarTxt, row);
                            if x.GetDate(Buffer, FIText, row) < Today then
                                Error(ErrorFecha, Today, Buffer."Row No.");

                            if SourceTypeText <> '' then
                                case
                                    SourceTypeText of
                                    'Cliente':
                                        SourceTypeEnum := SourceTypeEnum::Customer;
                                    'Grupo precio cliente':
                                        SourceTypeEnum := SourceTypeEnum::"Customer Price Group";
                                    'Todos los clientes':
                                        SourceTypeEnum := SourceTypeEnum::"All Customers";
                                    'Campaña':
                                        SourceTypeEnum := SourceTypeEnum::Campaign;
                                    'Grupo dto. cliente':
                                        SourceTypeEnum := SourceTypeEnum::"Customer Disc. Group";
                                    'Todos los proyectos':
                                        SourceTypeEnum := SourceTypeEnum::"All Jobs";
                                    'Proyecto':
                                        SourceTypeEnum := SourceTypeEnum::Job;
                                    'Tarea proyecto':
                                        SourceTypeEnum := SourceTypeEnum::"Job Task";
                                    'Contacto':
                                        SourceTypeEnum := SourceTypeEnum::Contact;
                                    else
                                        Error(ErrorOrtografia, Buffer."Row No.");
                                end;

                            if (SourceTypeEnum <> SourceTypeEnum::"All Customers") then begin
                                if x.GetAssignTo(buffer, NoAsignarTxt, row) = '' then
                                    Error(ErrorFaltaParametro, NoAsignarTxt, Buffer."Row No.");
                            end;


                            AmountTypeText := x.GetAmountType(Buffer, DefineTxt, row);
                            case
                                AmountTypeText of
                                'PRECIO Y DESCUENTO':
                                    AmountTypeEnum := AmountTypeEnum::Any;
                                'PRECIO':
                                    AmountTypeEnum := AmountTypeEnum::Price;
                                'DESCUENTO':
                                    AmountTypeEnum := AmountTypeEnum::Discount;
                            end;


                            PriceListLine.Reset();
                            PriceListLine.SetRange("Price List Code", rec.Code);
                            PriceListLine.SetRange("Source Type", SourceTypeEnum);
                            PriceListLine.SetRange("Assign-to No.", x.GetAssignTo(Buffer, NoAsignarTxt, row));
                            PriceListLine.SetRange("Product No.", x.GetNoProducto(Buffer, NoProductoTxt, row));
                            PriceListLine.SetRange("Amount Type", AmountTypeEnum);
                            PriceListLine.SetRange("Ending Date", 0D);
                            if PriceListLine.FindSet() then
                                if PriceListLine."Starting Date" = x.GetDate(Buffer, FIText, row) then
                                    //Si encuentra un record con la misma fecha, se le hara una correcion
                                    repeat
                                        //PriceListLine."Minimum Quantity" := x.GetCantidadMinima(Buffer, CantMinTxt, row);
                                        PriceListLine."Unit Price" := x.GetPrecioUnitario(Buffer, PrecioUnitarioTxt, row);
                                        PriceListLine."Line Discount %" := x.GetPorcientoLinea(buffer, PorcientoDscLinTxt, row);
                                        PriceListLine."DSLine Discount Amount" := x.GetImporteDescuento(buffer, ImpDescTxt, row);
                                        PriceListLine.Modify();
                                    until PriceListLine.Next() = 0

                                //si encuentra el mismo record con diferente fecha, le pone fecha fin
                                else

                                    if PriceListLine."Starting Date" > x.GetDate(Buffer, FIText, row) then
                                        Error(ErrorFecha, PriceListLine."Starting Date", Buffer."Row No.")
                                    else begin
                                        //Cambiar fecha final e insertar lineas
                                        PriceListLine."Ending Date" := x.GetDate(Buffer, FIText, row) - 1;
                                        PriceListLine.Modify();
                                        PriceListLine2.Reset();
                                        PriceListLine2.SetRange("Price List Code", rec.Code);
                                        PriceListLine2.FindLast();
                                        PriceListLine2."Line No." += 1;
                                        PriceListLine2."Source Type" := SourceTypeEnum;

                                        PriceListLine2."Assign-to No." := x.GetAssignTo(Buffer, NoAsignarTxt, row);
                                        PriceListLine2."Currency Code" := x.GetCodDivisa(Buffer, CodDivisaTxt, row);
                                        PriceListLine2."Starting Date" := x.GetDate(Buffer, FIText, row);
                                        PriceListLine2.Validate("Product No.", x.GetNoProducto(Buffer, NoProductoTxt, row));

                                        PriceListLine2."Unit of Measure Code" := x.GetCodUnidadMedida(Buffer, CodUnidadMedidaTxt, row);

                                        PriceListLine2."Amount Type" := AmountTypeEnum;
                                        PriceListLine2."Unit Price" := x.GetPrecioUnitario(Buffer, PrecioUnitarioTxt, row);
                                        PriceListLine2."Line Discount %" := x.GetPorcientoLinea(buffer, PorcientoDscLinTxt, Row);
                                        PriceListLine2."DSLine Discount Amount" := x.GetImporteDescuento(Buffer, ImpDescTxt, row);

                                        PriceListLine2."Ending Date" := 0D;

                                        if (PriceListLine2."Product No." = '') or (PriceListLine2."Unit of Measure Code" = '') or (PriceListLine2."Starting Date" = 0D) then
                                            Error(UploadErr)
                                        else
                                            PriceListLine2.Insert();
                                    end

                            //Insertar  linea si no existe record
                            else begin
                                PriceListLine2.Reset();
                                PriceListLine2.SetRange("Price List Code", rec.Code);
                                if PriceListLine2.FindLast() then begin
                                    PriceListLine2."Line No." += 1;
                                    PriceListLine2."Source Type" := SourceTypeEnum;

                                    PriceListLine2."Assign-to No." := x.GetAssignTo(Buffer, NoAsignarTxt, row);
                                    PriceListLine2."Currency Code" := x.GetCodDivisa(Buffer, CodDivisaTxt, row);
                                    PriceListLine2."Starting Date" := x.GetDate(Buffer, FIText, row);
                                    PriceListLine2.Validate("Product No.", x.GetNoProducto(Buffer, NoProductoTxt, row));

                                    PriceListLine2."Unit of Measure Code" := x.GetCodUnidadMedida(Buffer, CodUnidadMedidaTxt, row);

                                    PriceListLine2."Amount Type" := AmountTypeEnum;
                                    PriceListLine2."Unit Price" := x.GetPrecioUnitario(Buffer, PrecioUnitarioTxt, row);
                                    PriceListLine2."Line Discount %" := x.GetPorcientoLinea(buffer, PorcientoDscLinTxt, Row);
                                    PriceListLine2."DSLine Discount Amount" := x.GetImporteDescuento(Buffer, ImpDescTxt, row);


                                    if (PriceListLine2."Product No." = '') or (PriceListLine2."Unit of Measure Code" = '') or (PriceListLine2."Starting Date" = 0D) then
                                        Error(UploadErr)
                                    else
                                        PriceListLine2.Insert();
                                end
                                //Si no existe ninguna linea dentro de ese codigo de Price list
                                else begin
                                    PriceListLine2.Init();
                                    PriceListLine2."Price List Code" := rec.Code;
                                    PriceListLine2."Line No." := 1;
                                    PriceListLine2."Source Type" := SourceTypeEnum;
                                    PriceListLine2."Source No." := x.GetAssignTo(Buffer, NoAsignarTxt, row);
                                    PriceListLine2."Assign-to No." := x.GetAssignTo(Buffer, NoAsignarTxt, row);
                                    PriceListLine2."Currency Code" := x.GetCodDivisa(Buffer, CodDivisaTxt, row);
                                    PriceListLine2."Starting Date" := x.GetDate(Buffer, FIText, row);
                                    PriceListLine2.Validate("Product No.", x.GetNoProducto(Buffer, NoProductoTxt, row));

                                    PriceListLine2."Unit of Measure Code" := x.GetCodUnidadMedida(Buffer, CodUnidadMedidaTxt, row);

                                    PriceListLine2."Amount Type" := AmountTypeEnum;
                                    PriceListLine2."Unit Price" := x.GetPrecioUnitario(Buffer, PrecioUnitarioTxt, row);
                                    PriceListLine2."Line Discount %" := x.GetPorcientoLinea(buffer, PorcientoDscLinTxt, Row);
                                    PriceListLine2."DSLine Discount Amount" := x.GetImporteDescuento(Buffer, ImpDescTxt, row);


                                    if (PriceListLine2."Product No." = '') or (PriceListLine2."Unit of Measure Code" = '') or (PriceListLine2."Starting Date" = 0D) then
                                        Error(UploadErr)
                                    else
                                        PriceListLine2.Insert();
                                end;
                            end //FIN si no existe ninguna linea dentro de ese codigo de Price list
                        end
                    end
                end;

            }


        }
        addlast(Category_Process)
        {
            actionref(Importar_Promoted; ImportPrice)
            { }
        }
    }
    var
        Buffer: Record "Excel Buffer" temporary;

}
