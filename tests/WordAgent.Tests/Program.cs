using WordAgent.Core;

static void Assert(bool condition, string message)
{
    if (!condition) throw new Exception("SELFTEST FAILED: " + message);
}

static bool LoadRejected(string profilesDir, string profileId)
{
    try
    {
        _ = ProfileLoader.Load(profilesDir, profileId);
        return false;
    }
    catch (InvalidDataException)
    {
        return true;
    }
}

var root = Path.Combine(Path.GetTempPath(), "WordAgentSelfTest", Guid.NewGuid().ToString("N"));
Directory.CreateDirectory(root);
try
{
    var input = Path.Combine(root, "a.docx");
    File.WriteAllText(input, "x");
    var output = PathPolicy.CreateSafeOutputPath(input);
    Assert(!string.Equals(input, output, StringComparison.OrdinalIgnoreCase), "Safe output path must not overwrite input.");
    Assert(output.EndsWith("_WordAgent.docx", StringComparison.OrdinalIgnoreCase), "Output suffix is wrong.");

    var txt = Path.Combine(root, "a.txt");
    File.WriteAllText(txt, "x");
    var rejected = false;
    try { PathPolicy.RequireSupportedWordFile(txt); }
    catch (NotSupportedException) { rejected = true; }
    Assert(rejected, "Unsupported extension must be rejected.");

    var profiles = Path.Combine(root, "profiles");
    Directory.CreateDirectory(profiles);

    var legacyJson = """
    {
      "id": "legacy-profile",
      "displayName": "Legacy Profile",
      "version": "1.0.0",
      "documentClass": "LegacyTest",
      "language": "fa-IR",
      "classification": { "keywords": ["الف"], "negativeKeywords": [], "styleSignals": ["متن"] },
      "roleAliases": { "Body": ["متن", "Normal"] },
      "roleDefaults": { "Body": { "fontBi": "B Nazanin", "sizeBi": 12, "readingOrder": "RTL", "alignment": "RIGHT" } },
      "rules": ["FIELD-PRESERVE-001"],
      "referencePolicy": { "owner": "EndNote", "allowedModes": ["IEEE_STRICT"], "manualRenumber": false },
      "fieldPolicy": { "preserve": true, "updateTocOnRequestOnly": true },
      "outputPolicy": { "overwriteOriginal": false, "backupRequired": true }
    }
    """;
    File.WriteAllText(Path.Combine(profiles, "legacy-profile.json"), legacyJson);

    var legacy = ProfileLoader.Load(profiles, "legacy-profile");
    Assert(legacy.SchemaVersion == ProfileSchema.LegacyCompatibility,
        "Legacy profile must load through explicit compatibility schema.");
    Assert(legacy.DocumentClass == "LegacyTest", "Legacy metadata must deserialize.");
    Assert(legacy.Classification.StyleSignals.Contains("متن"), "Legacy classification must deserialize.");

    var v2Json = """
    {
      "schemaVersion": "2.0",
      "id": "v2-profile",
      "displayName": "V2 Profile",
      "version": "2.0.0",
      "documentClass": "V2Test",
      "language": "fa-IR",
      "classification": { "keywords": ["کتاب"], "negativeKeywords": [], "styleSignals": ["Heading 1"] },
      "roleAliases": { "Body": ["Normal"], "Heading1": ["Heading 1"] },
      "roleDefaults": {
        "Body": { "fontBi": "B Nazanin", "sizeBi": 12, "readingOrder": "RTL", "alignment": "RIGHT" },
        "Heading1": { "fontBi": "B Titr", "sizeBi": 16, "readingOrder": "RTL", "alignment": "RIGHT" }
      },
      "rules": ["FIELD-PRESERVE-001"],
      "referencePolicy": { "owner": "EndNote", "allowedModes": ["IEEE_STRICT"], "manualRenumber": false },
      "fieldPolicy": { "preserve": true, "updateTocOnRequestOnly": true },
      "outputPolicy": { "overwriteOriginal": false, "backupRequired": true }
    }
    """;
    File.WriteAllText(Path.Combine(profiles, "v2-profile.json"), v2Json);

    var v2 = ProfileLoader.Load(profiles, "v2-profile");
    Assert(v2.SchemaVersion == ProfileSchema.Current, "Explicit schema v2 must load.");
    Assert(ProfileValidator.Validate(v2).IsValid, "Valid v2 profile must pass validation.");

    var unsupportedSchema = v2Json.Replace(
        "\"schemaVersion\": \"2.0\"",
        "\"schemaVersion\": \"99.0\"");
    File.WriteAllText(Path.Combine(profiles, "bad-schema.json"), unsupportedSchema.Replace(
        "\"id\": \"v2-profile\"",
        "\"id\": \"bad-schema\""));
    Assert(LoadRejected(profiles, "bad-schema"), "Unsupported profile schema must be rejected.");

    var unsafeOverwrite = v2Json.Replace(
        "\"id\": \"v2-profile\"",
        "\"id\": \"unsafe-overwrite\"").Replace(
        "\"overwriteOriginal\": false",
        "\"overwriteOriginal\": true");
    File.WriteAllText(Path.Combine(profiles, "unsafe-overwrite.json"), unsafeOverwrite);
    Assert(LoadRejected(profiles, "unsafe-overwrite"), "Unsafe overwrite profile must be rejected.");

    var ambiguousAlias = v2Json.Replace(
        "\"id\": \"v2-profile\"",
        "\"id\": \"ambiguous-alias\"").Replace(
        "\"Heading1\": [\"Heading 1\"]",
        "\"Heading1\": [\"Normal\"]");
    File.WriteAllText(Path.Combine(profiles, "ambiguous-alias.json"), ambiguousAlias);
    Assert(LoadRejected(profiles, "ambiguous-alias"), "Ambiguous cross-role style alias must be rejected.");

    var invalidReadingOrder = v2Json.Replace(
        "\"id\": \"v2-profile\"",
        "\"id\": \"invalid-reading-order\"").Replace(
        "\"readingOrder\": \"RTL\"",
        "\"readingOrder\": \"SIDEWAYS\"");
    File.WriteAllText(Path.Combine(profiles, "invalid-reading-order.json"), invalidReadingOrder);
    Assert(LoadRejected(profiles, "invalid-reading-order"), "Invalid readingOrder must be rejected.");

    var semanticContract = new ProfileSemanticPolicy
    {
        TableRoles = new ProfileTableRoleSemanticPolicy
        {
            FirstRowRole = "TableHeader",
            BodyRole = "TableBody"
        },
        Invocation = new ProfileInvocationSemanticPolicy
        {
            Role = "Invocation",
            ExactCompactTexts = new[] { "بسمهتعالی", "بسماللهالرحمنالرحیم" }
        },
        CoverSubtitle = new ProfileCoverSubtitleSemanticPolicy
        {
            Role = "CoverSubtitle",
            StyleNames = new[] { "متن ضخیم" },
            MaxParagraphIndex = 20
        },
        FigureTableCaption = new ProfileFigureTableCaptionSemanticPolicy
        {
            Role = "FigureTableCaption",
            Prefixes = new[] { "جدول", "شکل" },
            RequireParenthesizedTwoPartNumber = true
        },
        Appendix = new ProfileAppendixSemanticPolicy
        {
            HeadingPrefixes = new[] { "پیوست" }
        }
    };

    Assert(
        SemanticPolicyMatcher.ResolveTableRole(1, semanticContract.TableRoles) == "TableHeader",
        "Semantic table first-row role must be profile-driven.");

    Assert(
        SemanticPolicyMatcher.ResolveTableRole(2, semanticContract.TableRoles) == "TableBody",
        "Semantic table body role must be profile-driven.");

    Assert(
        SemanticPolicyMatcher.MatchesInvocation("بسمه تعالی", semanticContract.Invocation),
        "Invocation matcher must compact profile-defined exact text.");

    Assert(
        SemanticPolicyMatcher.MatchesCoverSubtitle("متن ضخیم", 20, semanticContract.CoverSubtitle) &&
        !SemanticPolicyMatcher.MatchesCoverSubtitle("متن ضخیم", 21, semanticContract.CoverSubtitle),
        "Cover subtitle matcher must honor profile style names and max paragraph index.");

    Assert(
        SemanticPolicyMatcher.MatchesFigureTableCaption("جدول (۱-۲) نمونه", semanticContract.FigureTableCaption) &&
        !SemanticPolicyMatcher.MatchesFigureTableCaption("جدول نمونه", semanticContract.FigureTableCaption),
        "Caption matcher must honor profile prefixes and number requirement.");

    Assert(
        SemanticPolicyMatcher.MatchesAppendixHeading("پیوست الف", semanticContract.Appendix),
        "Appendix matcher must honor profile heading prefixes.");

    var semanticValidationProbe = ProfileLoader.Load(profiles, "v2-profile");
    semanticValidationProbe.SemanticPolicy = new ProfileSemanticPolicy
    {
        TableRoles = new ProfileTableRoleSemanticPolicy
        {
            FirstRowRole = "MissingRole"
        }
    };

    Assert(
        !ProfileValidator.Validate(semanticValidationProbe).IsValid,
        "Semantic policy must reject references to missing roleDefaults roles.");
    var layoutContract = new ProfileLayoutPolicy
    {
        Toc = new ProfileTocLayoutPolicy
        {
            Enabled = true,
            CompactStyles = true,
            TitleTexts = new[] { "فهرست مطالب" },
            Spacing = new ProfileTocSpacingLayoutPolicy
            {
                BeforeTwips = 0,
                AfterTwips = 0,
                LineTwips = 240,
                LineRule = "auto"
            },
            StartBodyOnNewPage = true
        },
        AppendixTables = new ProfileAppendixTableLayoutPolicy
        {
            Enabled = true,
            FitToTextWidth = true,
            RtlVisual = true,
            TableAlignment = "CENTER",
            FixedLayout = true,
            NormalizeCellMargins = true,
            CellMarginTwips = 70,
            WideDetection = new ProfileWideTableDetectionLayoutPolicy
            {
                MinColumns = 7,
                MinAverageColumnWidthTwips = 1200
            },
            LandscapeWideTables = new ProfileLandscapeWideTableLayoutPolicy
            {
                Enabled = true,
                Strategy = "FINAL_WIDE_TAIL_ONLY"
            }
        },
        TablePagination = new ProfileTablePaginationLayoutPolicy
        {
            Enabled = true,
            CompactMaxRows = 12,
            CompactScoreThreshold = 1050,
            RowPenaltyStart = 8,
            RowPenaltyWeight = 120,
            ColumnPenaltyStart = 6,
            ColumnPenaltyWeight = 100,
            NoRowSplit = true,
            RepeatHeader = true,
            KeepCaptionWithTable = true,
            MinBodyRowsAtStart = 2,
            MinBodyRowsAtEnd = 2
        }
    };

    Assert(
        LayoutPolicyDecisions.IsWideTable(7, 8400, layoutContract.AppendixTables.WideDetection),
        "Seven-column appendix table must satisfy profile wide-table threshold.");

    Assert(
        LayoutPolicyDecisions.IsWideTable(6, 6600, layoutContract.AppendixTables.WideDetection),
        "Narrow average column width must satisfy profile wide-table threshold.");

    Assert(
        !LayoutPolicyDecisions.IsWideTable(6, 7800, layoutContract.AppendixTables.WideDetection),
        "Non-wide table must remain portrait under profile thresholds.");

    Assert(
        LayoutPolicyDecisions.ShouldKeepTableWhole(12, 6, 100, layoutContract.TablePagination),
        "Compact pagination decision must preserve accepted dev7 behavior.");

    Assert(
        !LayoutPolicyDecisions.ShouldKeepTableWhole(13, 6, 100, layoutContract.TablePagination),
        "Row-count threshold must switch to controlled multi-page.");

    Assert(
        !LayoutPolicyDecisions.ShouldKeepTableWhole(12, 6, 700, layoutContract.TablePagination),
        "Profile score threshold must switch text-heavy table to controlled multi-page.");

    var layoutValidationProbe = ProfileLoader.Load(profiles, "v2-profile");
    layoutValidationProbe.LayoutPolicy = layoutContract;
    Assert(
        ProfileValidator.Validate(layoutValidationProbe).IsValid,
        "Explicit valid layout policy must pass validation.");

    layoutValidationProbe.LayoutPolicy.AppendixTables.CellMarginTwips = -1;
    Assert(
        !ProfileValidator.Validate(layoutValidationProbe).IsValid,
        "Negative profile cell margin must be rejected.");
    static void CreateDetectorFixture(string path, string text, params string[] styleNames)
    {
        using var fs = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None);
        using var zip = new System.IO.Compression.ZipArchive(
            fs,
            System.IO.Compression.ZipArchiveMode.Create,
            leaveOpen: false);

        static void AddZipEntry(
            System.IO.Compression.ZipArchive archive,
            string name,
            string content)
        {
            var entry = archive.CreateEntry(name);
            using var writer = new StreamWriter(
                entry.Open(),
                new System.Text.UTF8Encoding(false));
            writer.Write(content);
        }

        AddZipEntry(
            zip,
            "[Content_Types].xml",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/>" +
            "<Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/>" +
            "<Override PartName=\"/word/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml\"/>" +
            "</Types>");

        AddZipEntry(
            zip,
            "_rels/.rels",
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/>" +
            "</Relationships>");

        var stylesXml = new System.Text.StringBuilder();
        stylesXml.Append(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">");

        for (var i = 0; i < styleNames.Length; i++)
        {
            var escapedStyle = System.Security.SecurityElement.Escape(styleNames[i]) ?? string.Empty;
            stylesXml.Append(
                $"<w:style w:type=\"paragraph\" w:styleId=\"s{i}\"><w:name w:val=\"{escapedStyle}\"/></w:style>");
        }

        stylesXml.Append("</w:styles>");
        AddZipEntry(zip, "word/styles.xml", stylesXml.ToString());

        var escapedText = System.Security.SecurityElement.Escape(text) ?? string.Empty;
        var documentXml = new System.Text.StringBuilder();
        documentXml.Append(
            "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>" +
            "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");

        for (var i = 0; i < styleNames.Length; i++)
        {
            documentXml.Append(
                $"<w:p><w:pPr><w:pStyle w:val=\"s{i}\"/></w:pPr><w:r><w:t>{escapedText}</w:t></w:r></w:p>");
        }

        documentXml.Append("</w:body></w:document>");
        AddZipEntry(zip, "word/document.xml", documentXml.ToString());
    }
    var detectorProfiles = Path.Combine(Directory.GetCurrentDirectory(), "profiles");
    if (!Directory.Exists(detectorProfiles))
    {
        var testBase = AppContext.BaseDirectory;
        detectorProfiles = Path.Combine(testBase, "profiles");
    }

    if (Directory.Exists(detectorProfiles))
    {
        var detector = new DeterministicProfileDetector(detectorProfiles);

        var academicDoc = Path.Combine(root, "academic.docx");
        CreateDetectorFixture(
            academicDoc,
            "پیشنهاده رساله دکتری روش تحقیق مروری بر ادبیات اهداف تحقیق نوآوری پژوهش اساتید راهنما",
            "متن", "فصل", "تيتر اول", "عنوان پايان‌نامه", "زيرنويس شکل", "بالانويس جدول", "شکل - جدول");
        var academicResult = detector.DetectDetailedAsync(academicDoc).GetAwaiter().GetResult();
        Assert(academicResult.Status == ProfileDetectionStatus.Detected &&
               academicResult.ProfileId == "proposal.academic.fa",
            "Academic proposal detector fixture must resolve to proposal.academic.fa.");

        var phdDoc = Path.Combine(root, "phd.docx");
        CreateDetectorFixture(
            phdDoc,
            "رساله دکتری چکیده فصل مبانی نظری مرور ادبیات تحلیل نتایج نتیجه‌گیری دستاوردها مراجع",
            "متن", "فصل", "تيتر اول", "عنوان پايان‌نامه", "زيرنويس شکل", "بالانويس جدول", "شکل - جدول");
        var phdResult = detector.DetectDetailedAsync(phdDoc).GetAwaiter().GetResult();
        Assert(phdResult.Status == ProfileDetectionStatus.Detected &&
               phdResult.ProfileId == "thesis.dissertation.phd.fa",
            "PhD detector fixture must resolve to thesis.dissertation.phd.fa.");

        var bookDoc = Path.Combine(root, "book.docx");
        CreateDetectorFixture(
            bookDoc,
            "کتاب فصل بخش اهداف یادگیری تمرین مثال حل‌شده پیوست واژه‌نامه منابع و مآخذ راهنمای حل مسئله",
            "Title", "Heading 1", "Heading 2", "List Bullet", "List Number");
        var bookResult = detector.DetectDetailedAsync(bookDoc).GetAwaiter().GetResult();
        Assert(bookResult.Status == ProfileDetectionStatus.Detected &&
               bookResult.ProfileId == "book.university.fa",
            "Book detector fixture must resolve to book.university.fa.");

        var governanceDoc = Path.Combine(root, "governance.docx");
        CreateDetectorFixture(
            governanceDoc,
            "حاکمیت داده راهبری داده سیاست چارچوب کمیته مالک داده متولی داده مدل عملیاتی نقش و مسئولیت بلوغ کیفیت داده معماری داده",
            "متن", "فصل", "تيتر اول");
        var governanceResult = detector.DetectDetailedAsync(governanceDoc).GetAwaiter().GetResult();
        Assert(governanceResult.Status == ProfileDetectionStatus.Detected &&
               governanceResult.ProfileId == "proposal.governance.fa",
            "Governance detector fixture must resolve to proposal.governance.fa.");

        var industrialDoc = Path.Combine(root, "industrial.docx");
        CreateDetectorFixture(
            industrialDoc,
            "پروپوزال صنعتی پروژه صنعتی کارفرما شرح خدمات ساخت تجهیزات نمونه اولیه پیاده‌سازی تحویل‌دادنی بودجه زمان‌بندی آزمایش",
            "متن", "فصل", "تيتر اول");
        var industrialResult = detector.DetectDetailedAsync(industrialDoc).GetAwaiter().GetResult();
        Assert(industrialResult.Status == ProfileDetectionStatus.Detected &&
               industrialResult.ProfileId == "proposal.industrial.fa",
            "Industrial detector fixture must resolve to proposal.industrial.fa.");

        var questionnaireDoc = Path.Combine(root, "questionnaire.docx");
        CreateDetectorFixture(
            questionnaireDoc,
            "پرسشنامه امتیاز شاهد وضعیت موجود ردیف پاسخ ارزیابی",
            "متن", "تيتر اول");
        var questionnaireResult = detector.DetectDetailedAsync(questionnaireDoc).GetAwaiter().GetResult();
        Assert(questionnaireResult.Status == ProfileDetectionStatus.Detected &&
               questionnaireResult.ProfileId == "questionnaire-assessment-fa",
            "Questionnaire detector fixture must resolve to questionnaire-assessment-fa.");
    }
    Console.WriteLine("WordAgent self-tests: PASS");
    return 0;
}
finally
{
    try { Directory.Delete(root, true); } catch { }
}




