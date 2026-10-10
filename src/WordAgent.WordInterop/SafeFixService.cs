using Word = Microsoft.Office.Interop.Word;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed class SafeFixService
{
    public FixResult Apply(string inputPath, AgentProfile profile, string? outputDirectory = null, bool visible = false)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        var output = PathPolicy.CreateSafeOutputPath(input, outputDirectory);
        File.Copy(input, output, overwrite: false);

        var warnings = new List<string>();
        int stylesChanged = 0, tablesChanged = 0, tableParagraphsChanged = 0, inlineImagesCentered = 0;
        Word.Application? app = null;
        Word.Document? doc = null;
        WordAutomationSecurityScope? automationSecurity = null;

        try
        {
            app = new Word.Application();
            app.Visible = visible;
            app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            app.ScreenUpdating = visible;
            automationSecurity = WordAutomationSecurityScope.ForceDisable(app);

            doc = app.Documents.Open(
                FileName: output,
                ConfirmConversions: false,
                ReadOnly: false,
                AddToRecentFiles: false,
                Visible: visible,
                OpenAndRepair: false,
                NoEncodingDialog: true);

            var aliasMap = BuildAliasMap(profile);
            stylesChanged += UpdateStyles(doc, aliasMap, warnings);
            UpdateParagraphOverrides(doc, aliasMap, warnings);

            if (profile.RoleDefaults.TryGetValue("TableBody", out var tableDefaults))
            {
                (tablesChanged, tableParagraphsChanged) = UpdateTables(doc, tableDefaults, warnings);
            }

            inlineImagesCentered = CenterInlineImages(doc, warnings);

            doc.Save();
            return new FixResult(input, output, profile.Id, stylesChanged, tablesChanged, tableParagraphsChanged, inlineImagesCentered, warnings);
        }
        catch
        {
            try { if (File.Exists(output)) File.Delete(output); } catch { }
            throw;
        }
        finally
        {
            if (doc is not null)
            {
                try { doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                ComUtil.Release(doc);
            }
            automationSecurity?.Dispose();
            if (app is not null)
            {
                try { app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges); } catch { }
                ComUtil.Release(app);
            }
        }
    }

    private static Dictionary<string, RoleDefaults> BuildAliasMap(AgentProfile profile)
    {
        var map = new Dictionary<string, RoleDefaults>(StringComparer.OrdinalIgnoreCase);
        foreach (var (role, aliases) in profile.RoleAliases)
        {
            if (!profile.RoleDefaults.TryGetValue(role, out var defaults)) continue;
            foreach (var alias in aliases ?? Array.Empty<string>())
                if (!string.IsNullOrWhiteSpace(alias)) map[alias.Trim()] = defaults;
        }
        return map;
    }

    private static int UpdateStyles(Word.Document doc, IReadOnlyDictionary<string, RoleDefaults> aliasMap, List<string> warnings)
    {
        int changed = 0;
        Word.Styles? styles = null;
        try
        {
            styles = doc.Styles;
            for (int i = 1; i <= styles.Count; i++)
            {
                Word.Style? style = null;
                Word.Font? font = null;
                Word.ParagraphFormat? paragraphFormat = null;
                try
                {
                    style = styles[i];
                    var name = style.NameLocal;
                    if (!aliasMap.TryGetValue(name, out var defaults)) continue;

                    font = style.Font;
                    paragraphFormat = style.ParagraphFormat;
                    ApplyDefaults(font, paragraphFormat, defaults);
                    changed++;
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 80) warnings.Add($"Style {i}: {ex.Message}");
                }
                finally
                {
                    ComUtil.Release(paragraphFormat);
                    ComUtil.Release(font);
                    ComUtil.Release(style);
                }
            }
        }
        finally { ComUtil.Release(styles); }
        return changed;
    }

    private static void UpdateParagraphOverrides(Word.Document doc, IReadOnlyDictionary<string, RoleDefaults> aliasMap, List<string> warnings)
    {
        Word.Paragraphs? paragraphs = null;
        try
        {
            paragraphs = doc.Paragraphs;
            for (int i = 1; i <= paragraphs.Count; i++)
            {
                Word.Paragraph? p = null;
                Word.Range? range = null;
                Word.Style? style = null;
                Word.Font? font = null;
                Word.ParagraphFormat? format = null;
                try
                {
                    p = paragraphs[i];
                    range = p.Range;
                    if (ContainsProtectedEndNoteField(range)) continue;

                    style = range.get_Style() as Word.Style;
                    var styleName = style?.NameLocal;
                    if (string.IsNullOrWhiteSpace(styleName) || !aliasMap.TryGetValue(styleName, out var defaults)) continue;

                    font = range.Font;
                    format = range.ParagraphFormat;
                    ApplyDefaults(font, format, defaults);
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 80) warnings.Add($"Paragraph fix {i}: {ex.Message}");
                }
                finally
                {
                    ComUtil.Release(format);
                    ComUtil.Release(font);
                    ComUtil.Release(style);
                    ComUtil.Release(range);
                    ComUtil.Release(p);
                }
            }
        }
        finally { ComUtil.Release(paragraphs); }
    }

    private static (int TablesChanged, int ParagraphsChanged) UpdateTables(Word.Document doc, RoleDefaults defaults, List<string> warnings)
    {
        int tablesChanged = 0, paragraphsChanged = 0;
        Word.Tables? tables = null;
        try
        {
            tables = doc.Tables;
            for (int i = 1; i <= tables.Count; i++)
            {
                Word.Table? table = null;
                Word.Rows? rows = null;
                Word.Range? tableRange = null;
                Word.Paragraphs? paragraphs = null;
                try
                {
                    table = tables[i];
                    table.TableDirection = Word.WdTableDirection.wdTableDirectionRtl;
                    rows = table.Rows;
                    rows.Alignment = Word.WdRowAlignment.wdAlignRowCenter;
                    tablesChanged++;

                    tableRange = table.Range;
                    paragraphs = tableRange.Paragraphs;
                    for (int pIndex = 1; pIndex <= paragraphs.Count; pIndex++)
                    {
                        Word.Paragraph? p = null;
                        Word.Range? range = null;
                        Word.Font? font = null;
                        Word.ParagraphFormat? format = null;
                        try
                        {
                            p = paragraphs[pIndex];
                            range = p.Range;
                            if (ContainsProtectedEndNoteField(range)) continue;
                            font = range.Font;
                            format = range.ParagraphFormat;
                            ApplyDefaults(font, format, defaults);
                            paragraphsChanged++;
                        }
                        finally
                        {
                            ComUtil.Release(format);
                            ComUtil.Release(font);
                            ComUtil.Release(range);
                            ComUtil.Release(p);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 80) warnings.Add($"Table fix {i}: {ex.Message}");
                }
                finally
                {
                    ComUtil.Release(paragraphs);
                    ComUtil.Release(tableRange);
                    ComUtil.Release(rows);
                    ComUtil.Release(table);
                }
            }
        }
        finally { ComUtil.Release(tables); }
        return (tablesChanged, paragraphsChanged);
    }

    private static int CenterInlineImages(Word.Document doc, List<string> warnings)
    {
        int changed = 0;
        Word.InlineShapes? inlineShapes = null;
        try
        {
            inlineShapes = doc.InlineShapes;
            for (int i = 1; i <= inlineShapes.Count; i++)
            {
                Word.InlineShape? shape = null;
                Word.Range? range = null;
                Word.Paragraphs? paragraphs = null;
                Word.Paragraph? p = null;
                Word.Range? pRange = null;
                Word.ParagraphFormat? format = null;
                try
                {
                    shape = inlineShapes[i];
                    range = shape.Range;
                    paragraphs = range.Paragraphs;
                    if (paragraphs.Count < 1) continue;
                    p = paragraphs[1];
                    pRange = p.Range;
                    var text = WordAuditService.CleanParagraphText(pRange.Text);
                    if (!string.IsNullOrWhiteSpace(text)) continue;
                    format = pRange.ParagraphFormat;
                    format.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    changed++;
                }
                catch (Exception ex)
                {
                    if (warnings.Count < 80) warnings.Add($"Inline image {i}: {ex.Message}");
                }
                finally
                {
                    ComUtil.Release(format);
                    ComUtil.Release(pRange);
                    ComUtil.Release(p);
                    ComUtil.Release(paragraphs);
                    ComUtil.Release(range);
                    ComUtil.Release(shape);
                }
            }
        }
        finally { ComUtil.Release(inlineShapes); }
        return changed;
    }

    private static void ApplyDefaults(Word.Font font, Word.ParagraphFormat paragraphFormat, RoleDefaults defaults)
    {
        if (!string.IsNullOrWhiteSpace(defaults.FontBi))
            font.NameBi = defaults.FontBi;
        if (defaults.SizeBi is > 0)
            font.SizeBi = defaults.SizeBi.Value;

        if (string.Equals(defaults.ReadingOrder, "RTL", StringComparison.OrdinalIgnoreCase))
            paragraphFormat.ReadingOrder = Word.WdReadingOrder.wdReadingOrderRtl;
        else if (string.Equals(defaults.ReadingOrder, "LTR", StringComparison.OrdinalIgnoreCase))
            paragraphFormat.ReadingOrder = Word.WdReadingOrder.wdReadingOrderLtr;

        if (string.Equals(defaults.Alignment, "RIGHT", StringComparison.OrdinalIgnoreCase))
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
        else if (string.Equals(defaults.Alignment, "LEFT", StringComparison.OrdinalIgnoreCase))
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
        else if (string.Equals(defaults.Alignment, "CENTER", StringComparison.OrdinalIgnoreCase))
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
        else if (string.Equals(defaults.Alignment, "JUSTIFY", StringComparison.OrdinalIgnoreCase))
            paragraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
    }

    private static bool ContainsProtectedEndNoteField(Word.Range range)
    {
        Word.Fields? fields = null;
        try
        {
            fields = range.Fields;
            for (int i = 1; i <= fields.Count; i++)
            {
                Word.Field? field = null;
                Word.Range? code = null;
                try
                {
                    field = fields[i];
                    code = field.Code;
                    var text = code.Text ?? string.Empty;
                    if (text.Contains("EN.CITE", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("EN.BIB", StringComparison.OrdinalIgnoreCase) ||
                        text.Contains("EndNote", StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                finally { ComUtil.Release(code); ComUtil.Release(field); }
            }
            return false;
        }
        catch { return false; }
        finally { ComUtil.Release(fields); }
    }
}

