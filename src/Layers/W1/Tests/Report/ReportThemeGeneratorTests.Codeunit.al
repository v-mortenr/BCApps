// ------------------------------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT License. See License.txt in the project root for license information.
// ------------------------------------------------------------------------------------------------

codeunit 134620 "Report Theme Generator Tests"
{
    Subtype = Test;
    TestPermissions = Disabled;

    trigger OnRun()
    begin
        // [FEATURE] [Report Layout] [Composite Layout] [Report Theme Generator]
    end;

    var
        Assert: Codeunit Assert;
        LookupHelper: Codeunit "Composite Layout Lookup Helper";
        DefaultThemeTok: Label 'DefaultTheme', Locked = true;
        CalmThemeTok: Label 'CalmTheme', Locked = true;
        ThemeEntryTok: Label 'word/theme/theme1.xml', Locked = true;
        StylesEntryTok: Label 'word/styles.xml', Locked = true;
        DocumentEntryTok: Label 'word/document.xml', Locked = true;
        SettingsEntryTok: Label 'word/settings.xml', Locked = true;

    [Test]
    procedure ColorReplacesSystemColorWithRgb()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] A colour slot that holds a:sysClr in the base theme ends up as a:srgbClr with the new value.
        // [GIVEN] The shipped Default theme, whose Light 1 slot is a:sysClr.
        ReportThemeGenerator.SetColor("Report Theme Color Slot"::Light1, '#fff8e7');

        // [WHEN] The theme is generated.
        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        // [THEN] Light 1 is a normalized RGB value and the system colour is gone.
        Assert.AreEqual('FFF8E7', GetValue(GeneratedTheme, ThemeEntryTok, '//a:clrScheme/a:lt1/a:srgbClr/@val'), 'Light 1');
        Assert.IsFalse(HasNode(GeneratedTheme, ThemeEntryTok, '//a:clrScheme/a:lt1/a:sysClr'), 'Light 1 still holds a system colour');
    end;

    [Test]
    procedure UnsetColorsKeepBaseValues()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        BaseTheme: Codeunit "Temp Blob";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] Only the slots that were set change.
        GetShippedTheme(DefaultThemeTok, BaseTheme);
        ReportThemeGenerator.SetColor("Report Theme Color Slot"::Accent1, '4D614C');

        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        Assert.AreEqual('4D614C', GetValue(GeneratedTheme, ThemeEntryTok, '//a:clrScheme/a:accent1/a:srgbClr/@val'), 'Accent 1');
        Assert.AreEqual(
            GetValue(BaseTheme, ThemeEntryTok, '//a:clrScheme/a:accent3/a:srgbClr/@val'),
            GetValue(GeneratedTheme, ThemeEntryTok, '//a:clrScheme/a:accent3/a:srgbClr/@val'),
            'Accent 3 must keep the base value');
    end;

    [Test]
    procedure FontsAreWrittenToFontScheme()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] Heading and body fonts land in the major and minor fonts of the theme's font scheme.
        ReportThemeGenerator.SetFonts('Georgia', 'Verdana');

        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        Assert.AreEqual('Georgia', GetValue(GeneratedTheme, ThemeEntryTok, '//a:fontScheme/a:majorFont/a:latin/@typeface'), 'Heading font');
        Assert.AreEqual('Verdana', GetValue(GeneratedTheme, ThemeEntryTok, '//a:fontScheme/a:minorFont/a:latin/@typeface'), 'Body font');
    end;

    [Test]
    procedure ThemeNameIsWrittenToSchemes()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] The theme name is shown in Word's Design tab, so it is written to the theme and both schemes.
        ReportThemeGenerator.SetThemeName('Nature');

        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        Assert.AreEqual('Nature', GetValue(GeneratedTheme, ThemeEntryTok, '/a:theme/@name'), 'Theme name');
        Assert.AreEqual('Nature', GetValue(GeneratedTheme, ThemeEntryTok, '//a:clrScheme/@name'), 'Colour scheme name');
        Assert.AreEqual('Nature', GetValue(GeneratedTheme, ThemeEntryTok, '//a:fontScheme/@name'), 'Font scheme name');
    end;

    [Test]
    procedure PageColorIsAddedWithDisplayFlag()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] A page colour on a theme without one adds w:background first in w:document and turns on
        // w:displayBackgroundShape so Word shows it while editing.
        // [GIVEN] The Default theme, which has neither.
        ReportThemeGenerator.SetPageColor('EAEDE4');

        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        Assert.AreEqual('EAEDE4', GetValue(GeneratedTheme, DocumentEntryTok, '/w:document/w:background/@w:color'), 'Page colour');
        Assert.IsTrue(HasNode(GeneratedTheme, DocumentEntryTok, '/w:document/*[1][self::w:background]'), 'w:background must be the first child of w:document');
        Assert.IsTrue(HasNode(GeneratedTheme, SettingsEntryTok, '/w:settings/w:displayBackgroundShape'), 'displayBackgroundShape');
    end;

    [Test]
    procedure EmptyPageColorRemovesBackground()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] Clearing the page colour removes the base theme's background.
        // [GIVEN] The Calm theme, which has a page colour.
        ReportThemeGenerator.SetPageColor('');

        Generate(ReportThemeGenerator, CalmThemeTok, GeneratedTheme);

        Assert.IsFalse(HasNode(GeneratedTheme, DocumentEntryTok, '/w:document/w:background'), 'Page colour was not removed');
    end;

    [Test]
    procedure BaseFontSizeIsWrittenToDocDefaults()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
    begin
        // [SCENARIO] The base font size is stored in half-points in the docDefaults of styles.xml, which compose keeps from the theme.
        ReportThemeGenerator.SetBaseFontSize(10.5);

        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        Assert.AreEqual('21', GetValue(GeneratedTheme, StylesEntryTok, '/w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val'), 'sz');
        Assert.AreEqual('21', GetValue(GeneratedTheme, StylesEntryTok, '/w:styles/w:docDefaults/w:rPrDefault/w:rPr/w:szCs/@w:val'), 'szCs');
    end;

    [Test]
    procedure InvalidColorIsRejected()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
    begin
        // [SCENARIO] Only six-digit hex colours are accepted.
        asserterror ReportThemeGenerator.SetColor("Report Theme Color Slot"::Accent1, 'green');
        Assert.ExpectedError('is not a valid colour');

        asserterror ReportThemeGenerator.SetPageColor('12345G');
        Assert.ExpectedError('is not a valid colour');
    end;

    [Test]
    procedure NonThemeFileIsRejected()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        DataCompression: Codeunit "Data Compression";
        NotATheme: Codeunit "Temp Blob";
        GeneratedTheme: Codeunit "Temp Blob";
        ContentOutStream: OutStream;
    begin
        // [SCENARIO] A package without a theme part is refused with a clear error.
        DataCompression.CreateZipArchive();
        NotATheme.CreateOutStream(ContentOutStream);
        ContentOutStream.WriteText('<w:document/>');
        DataCompression.AddEntry(NotATheme.CreateInStream(), DocumentEntryTok);
        DataCompression.SaveZipArchive(NotATheme);

        asserterror ReportThemeGenerator.Generate(NotATheme.CreateInStream(), GeneratedTheme);
        Assert.ExpectedError(ThemeEntryTok);
    end;

    [Test]
    procedure GeneratedThemePassesLayoutValidation()
    var
        ReportThemeGenerator: Codeunit "Report Theme Generator";
        GeneratedTheme: Codeunit "Temp Blob";
        PreparedTheme: Codeunit "Temp Blob";
        PreparedInStream: InStream;
    begin
        // [SCENARIO] The generated file is still a layout the platform accepts for a theme part.
        ReportThemeGenerator.SetThemeName('Nature');
        ReportThemeGenerator.SetColor("Report Theme Color Slot"::Accent1, '4D614C');
        ReportThemeGenerator.SetFonts('Georgia', 'Verdana');
        ReportThemeGenerator.SetBaseFontSize(9);
        ReportThemeGenerator.SetPageColor('F7F7F2');
        Generate(ReportThemeGenerator, DefaultThemeTok, GeneratedTheme);

        // The platform rejects layout streams that are not declared UTF-8.
        PreparedInStream := PreparedTheme.CreateInStream(TextEncoding::UTF8);
        Assert.IsTrue(
            Report.ValidateAndPrepareLayout(LookupHelper.GetTenantReportDefaultsReportID(), GeneratedTheme.CreateInStream(TextEncoding::UTF8), PreparedInStream, ReportLayoutType::Word),
            'The generated theme was rejected by layout validation: ' + GetLastErrorText());
    end;

    local procedure Generate(var ReportThemeGenerator: Codeunit "Report Theme Generator"; BaseThemeName: Text; var GeneratedTheme: Codeunit "Temp Blob")
    var
        BaseTheme: Codeunit "Temp Blob";
    begin
        GetShippedTheme(BaseThemeName, BaseTheme);
        ReportThemeGenerator.Generate(BaseTheme.CreateInStream(), GeneratedTheme);
    end;

    local procedure GetShippedTheme(LayoutName: Text; var Theme: Codeunit "Temp Blob")
    var
        ReportLayoutList: Record "Report Layout List";
        LayoutOutStream: OutStream;
    begin
        ReportLayoutList.SetRange("Report ID", LookupHelper.GetTenantReportDefaultsReportID());
        ReportLayoutList.SetRange(Name, LayoutName);
        ReportLayoutList.SetRange("Layout Subtype", ReportLayoutList."Layout Subtype"::Theme);
        ReportLayoutList.FindFirst();
        Theme.CreateOutStream(LayoutOutStream);
        ReportLayoutList.Layout.ExportStream(LayoutOutStream);
    end;

    local procedure LoadPart(Package: Codeunit "Temp Blob"; EntryName: Text; var Document: XmlDocument; var NamespaceManager: XmlNamespaceManager)
    var
        DataCompression: Codeunit "Data Compression";
        Entry: Codeunit "Temp Blob";
    begin
        DataCompression.OpenZipArchive(Package, false);
        DataCompression.ExtractEntry(EntryName, Entry);
        DataCompression.CloseZipArchive();
        XmlDocument.ReadFrom(Entry.CreateInStream(TextEncoding::UTF8), Document);
        NamespaceManager.NameTable(Document.NameTable());
        NamespaceManager.AddNamespace('w', 'http://schemas.openxmlformats.org/wordprocessingml/2006/main');
        NamespaceManager.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main');
    end;

    local procedure GetValue(Package: Codeunit "Temp Blob"; EntryName: Text; XPath: Text): Text
    var
        Document: XmlDocument;
        NamespaceManager: XmlNamespaceManager;
        Node: XmlNode;
    begin
        LoadPart(Package, EntryName, Document, NamespaceManager);
        Assert.IsTrue(Document.SelectSingleNode(XPath, NamespaceManager, Node), StrSubstNo('%1 not found in %2', XPath, EntryName));
        exit(Node.AsXmlAttribute().Value());
    end;

    local procedure HasNode(Package: Codeunit "Temp Blob"; EntryName: Text; XPath: Text): Boolean
    var
        Document: XmlDocument;
        NamespaceManager: XmlNamespaceManager;
        Node: XmlNode;
    begin
        LoadPart(Package, EntryName, Document, NamespaceManager);
        exit(Document.SelectSingleNode(XPath, NamespaceManager, Node));
    end;
}
