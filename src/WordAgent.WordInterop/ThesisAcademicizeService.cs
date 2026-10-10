using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using WordAgent.Core;

namespace WordAgent.WordInterop;

public sealed record ThesisAcademicizeResult(
    string InputFile,
    string OutputFile,
    string InputSha256,
    string OutputSha256,
    string SelectedProfileId,
    string RoutePolicy,
    int HeadingParagraphsRightAligned,
    int BodyParagraphsJustified,
    int Tables,
    int TableCaptionsInserted,
    int Drawings,
    int FigureCaptionsInserted,
    int ChartCaptionsInserted,
    bool TocInserted,
    bool OriginalOverwritten);

public sealed class ThesisAcademicizeService
{
    // R10R39: schema-order normalization for w:bidi/w:jc/w:rtl on thesis-generated OOXML.
    private static readonly XNamespace W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public ThesisAcademicizeResult Apply(string inputPath, string selectedProfileId, string? outputDirectory = null)
    {
        var input = PathPolicy.RequireSupportedWordFile(inputPath);
        var normalizedProfileId = (selectedProfileId ?? string.Empty).Trim();
        var isGenericThesis = normalizedProfileId.Equals(
            "thesis-dissertation-academic-fa",
            StringComparison.OrdinalIgnoreCase);
        var isPhdThesis = normalizedProfileId.Equals(
            "thesis.dissertation.phd.fa",
            StringComparison.OrdinalIgnoreCase);
        if (!isGenericThesis && !isPhdThesis)
            throw new InvalidOperationException("Unsupported specialized thesis profile: " + normalizedProfileId);

        var routePolicy = isGenericThesis
            ? "GENERIC_ACADEMIC_THESIS"
            : "PHD_DISSERTATION";
        var routeSuffix = isGenericThesis
            ? "_WordAgent_Thesis_Academic"
            : "_WordAgent_Thesis_PhD";

        var outDir = string.IsNullOrWhiteSpace(outputDirectory) ? Path.GetDirectoryName(input)! : outputDirectory!;
        Directory.CreateDirectory(outDir);
        var stem = Path.GetFileNameWithoutExtension(input);
        var output = Path.Combine(outDir, stem + routeSuffix + ".docx");
        if (File.Exists(output))
            output = Path.Combine(outDir, stem + routeSuffix + "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss", CultureInfo.InvariantCulture) + ".docx");
        File.Copy(input, output, false);

        var inputSha = Sha(input);
        int headingRight=0, bodyJustified=0, tables=0, tableCaps=0, drawings=0, figCaps=0, chartCaps=0;
        bool tocInserted=false;

        using (var archive = ZipFile.Open(output, ZipArchiveMode.Update))
        {
            var docEntry = archive.GetEntry("word/document.xml") ?? throw new InvalidOperationException("word/document.xml missing.");
            XDocument doc;
            using (var s=docEntry.Open()) doc=XDocument.Load(s, LoadOptions.PreserveWhitespace);

            var styles = new Dictionary<string,string>(StringComparer.OrdinalIgnoreCase);
            var stylesEntry=archive.GetEntry("word/styles.xml");
            if (stylesEntry is not null)
            {
                using var ss=stylesEntry.Open();
                var sx=XDocument.Load(ss);
                foreach (var st in sx.Descendants(W+"style"))
                {
                    var id=((string?)st.Attribute(W+"styleId")??"").Trim();
                    var nm=((string?)st.Element(W+"name")?.Attribute(W+"val")??"").Trim();
                    if(id.Length>0) styles[id]=nm;
                }
            }

            var body=doc.Root?.Element(W+"body") ?? throw new InvalidOperationException("document body missing.");
            var blocks=body.Elements().ToList();

            foreach(var p in body.Descendants(W+"p").ToList())
            {
                var text=string.Concat(p.Descendants(W+"t").Select(t=>(string?)t??"")).Trim();
                if(!ContainsPersian(text)) continue;
                var styleId=((string?)p.Element(W+"pPr")?.Element(W+"pStyle")?.Attribute(W+"val")??"").Trim();
                styles.TryGetValue(styleId,out var styleName);
                var key=((styleId??"")+"|"+(styleName??"")).Replace(" ","").Replace("-","").Replace("_","").ToLowerInvariant();
                var isHeading=key.Contains("heading1")||key.Contains("heading2")||key.Contains("heading3")||key.Contains("تیتر")||key.Contains("تيتر");
                var isTocTitle=IsExactTocTitle(text);
                var isMainTocField=HasMainTocField(p);
                var isTitle=key.Contains("title")||key.Contains("عنوانپایان")||key.Contains("عنوانپايان")||key.Contains("chapter")||key.Contains("فصل")||isTocTitle;
                var isCaption=LooksCaption(text);
                var pPr=EnsurePPr(p);
                EnsureOn(pPr,"bidi");
                if(isHeading)
                {
                    SetJc(pPr,"right"); headingRight++;
                }
                else if(isTocTitle)
                {
                    SetJc(pPr,"center");
                    MaterializeRunContract(p,"B Nazanin",32);
                }
                else if(isMainTocField)
                {
                    SetJc(pPr,"start");
                    MaterializeRunContract(p,"B Nazanin",24);
                }
                else if(!isTitle && !isCaption && !p.Ancestors(W+"tbl").Any())
                {
                    SetJc(pPr,"both");
                    MaterializeRunContract(p,"B Nazanin",28);
                    bodyJustified++;
                }
                foreach(var r in p.Descendants(W+"r"))
                {
                    var rt=string.Concat(r.Descendants(W+"t").Select(t=>(string?)t??""));
                    if(!ContainsPersian(rt)) continue;
                    var rPr=r.Element(W+"rPr");
                    if(rPr is null){rPr=new XElement(W+"rPr");r.AddFirst(rPr);}
                    EnsureOn(rPr,"rtl");
                }
            }

            // Insert TOC after first explicit page-break paragraph, otherwise at body start.
            if(!body.Descendants(W+"instrText").Any(x=>((string?)x??"").Contains("TOC",StringComparison.OrdinalIgnoreCase)))
            {
                var title=NewParagraph("فهرست مطالب","center",bold:true);
                MaterializeRunContract(title,"B Nazanin",32);
                var toc=new XElement(W+"p",
                    new XElement(W+"pPr",new XElement(W+"bidi"),new XElement(W+"jc",new XAttribute(W+"val","start"))),
                    new XElement(W+"r",new XElement(W+"fldChar",new XAttribute(W+"fldCharType","begin"))),
                    new XElement(W+"r",new XElement(W+"instrText",new XAttribute(XNamespace.Xml+"space","preserve")," TOC \\o \"1-3\" \\h \\z \\u ")),
                    new XElement(W+"r",new XElement(W+"fldChar",new XAttribute(W+"fldCharType","separate"))),
                    new XElement(W+"r",new XElement(W+"t","برای نمایش فهرست، در Word روی Update Field کلیک کنید.")),
                    new XElement(W+"r",new XElement(W+"fldChar",new XAttribute(W+"fldCharType","end"))));
                MaterializeRunContract(toc,"B Nazanin",24);
                var br=NewPageBreakParagraph();
                var firstBreak=body.Elements(W+"p").FirstOrDefault(p=>p.Descendants(W+"br").Any(b=>((string?)b.Attribute(W+"type")??"").Equals("page",StringComparison.OrdinalIgnoreCase)));
                if(firstBreak is not null){firstBreak.AddAfterSelf(title);title.AddAfterSelf(toc);toc.AddAfterSelf(br);}
                else {var first=body.Elements().FirstOrDefault(); if(first is not null){first.AddBeforeSelf(title);title.AddAfterSelf(toc);toc.AddAfterSelf(br);} else {body.Add(title,toc,br);}}
                tocInserted=true;
            }

            // Existing table captions belong above the table and are centered.
            // Do not synthesize placeholder captions in the release candidate.
            // Missing-caption generation is intentionally deferred to the later Smart Caption phase.
            foreach(var tbl in body.Elements(W+"tbl").ToList())
            {
                tables++;
                var prev=tbl.ElementsBeforeSelf().LastOrDefault();
                if(prev?.Name==W+"p" && LooksTableCaption(TextOf(prev)))
                {
                    var pp=EnsurePPr(prev); EnsureOn(pp,"bidi"); SetJc(pp,"center");
                    tableCaps++;
                    continue;
                }

                var next=tbl.ElementsAfterSelf().FirstOrDefault();
                if(next?.Name==W+"p" && LooksTableCaption(TextOf(next)))
                {
                    next.Remove();
                    tbl.AddBeforeSelf(next);
                    var pp=EnsurePPr(next); EnsureOn(pp,"bidi"); SetJc(pp,"center");
                    tableCaps++;
                }
            }

            // Existing figure/chart captions belong below the drawing and are centered.
            // Release candidate must not invent bare captions such as "شکل ۱" for cover/logo/front-matter images.
            foreach(var p in body.Elements(W+"p").Where(p=>p.Descendants(W+"drawing").Any()).ToList())
            {
                drawings++;
                var next=p.ElementsAfterSelf().FirstOrDefault();
                if(next?.Name==W+"p" && LooksFigureCaption(TextOf(next)))
                {
                    var pp=EnsurePPr(next); EnsureOn(pp,"bidi"); SetJc(pp,"center");
                    if(TextOf(next).TrimStart().StartsWith("نمودار",StringComparison.Ordinal)) chartCaps++; else figCaps++;
                    continue;
                }

                var prev=p.ElementsBeforeSelf().LastOrDefault();
                if(prev?.Name==W+"p" && LooksFigureCaption(TextOf(prev)))
                {
                    var existing=prev; existing.Remove(); p.AddAfterSelf(existing);
                    var pp=EnsurePPr(existing); EnsureOn(pp,"bidi"); SetJc(pp,"center");
                    if(TextOf(existing).TrimStart().StartsWith("نمودار",StringComparison.Ordinal)) chartCaps++; else figCaps++;
                }
            }

            ReplaceEntry(archive,"word/document.xml",doc);

            // Ask Word to refresh fields on open.
            var settingsEntry=archive.GetEntry("word/settings.xml");
            if(settingsEntry is not null)
            {
                XDocument sx; using(var s=settingsEntry.Open()) sx=XDocument.Load(s,LoadOptions.PreserveWhitespace);
                var root=sx.Root!;
                var uf=root.Element(W+"updateFields");
                if(uf is null){uf=new XElement(W+"updateFields");root.Add(uf);}
                uf.SetAttributeValue(W+"val","true");
                ReplaceEntry(archive,"word/settings.xml",sx);
            }
        }

        _ = new PersianDocumentFoundationService().ApplyInPlace(
            output,
            new PersianDocumentFoundationOptions(ProfileId: normalizedProfileId));

        return new ThesisAcademicizeResult(input,output,inputSha,Sha(output),normalizedProfileId,routePolicy,headingRight,bodyJustified,tables,tableCaps,drawings,figCaps,chartCaps,tocInserted,false);
    }

    private static string TextOf(XElement p)=>string.Concat(p.Descendants(W+"t").Select(t=>(string?)t??"")).Trim();
    private static bool ContainsPersian(string s)=>s.Any(ch=>(ch>='\u0600'&&ch<='\u06FF'));
    private static bool LooksCaption(string s)=>LooksTableCaption(s)||LooksFigureCaption(s);
    private static bool LooksTableCaption(string s)=>s.TrimStart().StartsWith("جدول",StringComparison.Ordinal);
    private static bool LooksFigureCaption(string s){var t=s.TrimStart();return t.StartsWith("شکل",StringComparison.Ordinal)||t.StartsWith("نمودار",StringComparison.Ordinal);}
    private static bool HasField(XElement p)=>p.Descendants(W+"instrText").Any()||p.Descendants(W+"fldChar").Any();
    private static XElement EnsurePPr(XElement p){var x=p.Element(W+"pPr");if(x is null){x=new XElement(W+"pPr");p.AddFirst(x);}return x;}
    private static void EnsureOn(XElement parent,string name)
    {
        var x=parent.Element(W+name);
        if(x is null)x=new XElement(W+name);
        else { x.Remove(); x.SetAttributeValue(W+"val","1"); }
        if(parent.Name==W+"pPr")InsertParagraphPropertyInSchemaOrder(parent,x);
        else if(parent.Name==W+"rPr")InsertRunPropertyInSchemaOrder(parent,x);
        else parent.Add(x);
    }
    private static void SetJc(XElement pPr,string val)
    {
        var jc=pPr.Element(W+"jc");
        if(jc is null)jc=new XElement(W+"jc");else jc.Remove();
        jc.SetAttributeValue(W+"val",val);
        InsertParagraphPropertyInSchemaOrder(pPr,jc);
    }
    private static void InsertParagraphPropertyInSchemaOrder(XElement pPr,XElement property)
    {
        static int Rank(string name)=>name switch
        {
            "pStyle"=>10,"keepNext"=>20,"keepLines"=>30,"pageBreakBefore"=>40,"framePr"=>50,
            "widowControl"=>60,"numPr"=>70,"suppressLineNumbers"=>80,"pBdr"=>90,"shd"=>100,
            "tabs"=>110,"suppressAutoHyphens"=>120,"kinsoku"=>130,"wordWrap"=>140,"overflowPunct"=>150,
            "topLinePunct"=>160,"autoSpaceDE"=>170,"autoSpaceDN"=>180,"bidi"=>190,"adjustRightInd"=>200,
            "snapToGrid"=>210,"spacing"=>220,"ind"=>230,"contextualSpacing"=>240,"mirrorIndents"=>250,
            "suppressOverlap"=>260,"jc"=>270,"textDirection"=>280,"textAlignment"=>290,"textboxTightWrap"=>300,
            "outlineLvl"=>310,"divId"=>320,"cnfStyle"=>330,"rPr"=>340,"sectPr"=>350,"pPrChange"=>360,_=>1000
        };
        var target=Rank(property.Name.LocalName);
        var next=pPr.Elements().FirstOrDefault(e=>Rank(e.Name.LocalName)>target);
        if(next is null)pPr.Add(property);else next.AddBeforeSelf(property);
    }
    private static void InsertRunPropertyInSchemaOrder(XElement rPr,XElement property)
    {
        static int Rank(string name)=>name switch
        {
            "rStyle"=>10,"rFonts"=>20,"b"=>30,"bCs"=>40,"i"=>50,"iCs"=>60,"caps"=>70,"smallCaps"=>80,
            "strike"=>90,"dstrike"=>100,"outline"=>110,"shadow"=>120,"emboss"=>130,"imprint"=>140,"noProof"=>150,
            "snapToGrid"=>160,"vanish"=>170,"webHidden"=>180,"color"=>190,"spacing"=>200,"w"=>210,"kern"=>220,
            "position"=>230,"sz"=>240,"szCs"=>250,"highlight"=>260,"u"=>270,"effect"=>280,"bdr"=>290,"shd"=>300,
            "fitText"=>310,"vertAlign"=>320,"rtl"=>330,"cs"=>340,"em"=>350,"lang"=>360,"eastAsianLayout"=>370,
            "specVanish"=>380,"oMath"=>390,"rPrChange"=>400,_=>1000
        };
        var target=Rank(property.Name.LocalName);
        var next=rPr.Elements().FirstOrDefault(e=>Rank(e.Name.LocalName)>target);
        if(next is null)rPr.Add(property);else next.AddBeforeSelf(property);
    }
    private static bool IsExactTocTitle(string text)
        => string.Equals(
            string.Concat((text ?? string.Empty).Where(ch => !char.IsWhiteSpace(ch))),
            "فهرستمطالب",
            StringComparison.Ordinal);

    private static bool HasMainTocField(XElement paragraph)
    {
        var instruction=string.Join(
            " ",
            paragraph.Descendants(W+"instrText").Select(x=>((string?)x??string.Empty).Trim()));
        var normalized=string.Join(
            " ",
            instruction.Split((char[]?)null,StringSplitOptions.RemoveEmptyEntries));
        return normalized.StartsWith("TOC",StringComparison.OrdinalIgnoreCase)
               && (normalized.Length==3 || char.IsWhiteSpace(normalized[3]));
    }

    private static void MaterializeRunContract(XElement paragraph,string fontBi,int sizeHalfPoints)
    {
        var pPr=EnsurePPr(paragraph);
        var pMark=pPr.Element(W+"rPr");
        if(pMark is null){pMark=new XElement(W+"rPr");InsertParagraphPropertyInSchemaOrder(pPr,pMark);}
        SetRunContract(pMark,fontBi,sizeHalfPoints);
        foreach(var run in paragraph.Elements(W+"r"))
        {
            var rPr=run.Element(W+"rPr");
            if(rPr is null){rPr=new XElement(W+"rPr");run.AddFirst(rPr);}
            SetRunContract(rPr,fontBi,sizeHalfPoints);
            EnsureOn(rPr,"rtl");
        }
    }

    private static void SetRunContract(XElement rPr,string fontBi,int sizeHalfPoints)
    {
        var rf=rPr.Element(W+"rFonts");
        if(rf is null){rf=new XElement(W+"rFonts");InsertRunPropertyInSchemaOrder(rPr,rf);}
        rf.SetAttributeValue(W+"cs",fontBi);
        var sz=rPr.Element(W+"szCs");
        if(sz is null){sz=new XElement(W+"szCs");InsertRunPropertyInSchemaOrder(rPr,sz);}
        sz.SetAttributeValue(W+"val",sizeHalfPoints.ToString(CultureInfo.InvariantCulture));
    }

    private static XElement NewParagraph(string text,string align,bool bold,string? styleId=null)
    {
        var pPr=new XElement(W+"pPr");
        if(!string.IsNullOrWhiteSpace(styleId))InsertParagraphPropertyInSchemaOrder(pPr,new XElement(W+"pStyle",new XAttribute(W+"val",styleId)));
        EnsureOn(pPr,"bidi");
        SetJc(pPr,align);
        var rPr=new XElement(W+"rPr");
        EnsureOn(rPr,"rtl");
        if(bold)InsertRunPropertyInSchemaOrder(rPr,new XElement(W+"b"));
        return new XElement(W+"p",pPr,new XElement(W+"r",rPr,new XElement(W+"t",text)));
    }
    private static XElement NewPageBreakParagraph()=>new XElement(W+"p",new XElement(W+"r",new XElement(W+"br",new XAttribute(W+"type","page"))));
    private static string ToPersian(int n)=>n.ToString(CultureInfo.InvariantCulture).TranslateDigits();
    private static void ReplaceEntry(ZipArchive archive,string name,XDocument doc)
    {
        archive.GetEntry(name)?.Delete();var e=archive.CreateEntry(name,CompressionLevel.Optimal);
        using var s=e.Open();using var w=new StreamWriter(s,new UTF8Encoding(false));doc.Save(w,SaveOptions.DisableFormatting);
    }
    private static string Sha(string p){using var s=File.OpenRead(p);return Convert.ToHexString(SHA256.HashData(s));}
}

internal static class ThesisDigitExtensions
{
    public static string TranslateDigits(this string value)
    {
        const string en="0123456789";const string fa="۰۱۲۳۴۵۶۷۸۹";
        var chars=value.ToCharArray();
        for(int i=0;i<chars.Length;i++){var j=en.IndexOf(chars[i]);if(j>=0)chars[i]=fa[j];}
        return new string(chars);
    }
}
