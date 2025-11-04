using System;
using System.IO;
using PdfSharpCore.Fonts;

public class KoreanFontResolver : IFontResolver
{
    // 문서에서 쓸 이름 (마음대로, 스타일에서 이 이름을 참조)
    public const string FamilyName = "MalgunGothicEmbedded";

    // 실제 TTF 경로 (배포 시 폰트 동봉이나 상대경로로 바꾸세요)
    private const string RegularPath = @"C:\Windows\Fonts\malgun.ttf";
    private const string BoldPath = @"C:\Windows\Fonts\malgunbd.ttf";

    public string DefaultFontName => FamilyName;

    public byte[] GetFont(string faceName)
    {
        string path = faceName switch
        {
            FamilyName + "#b" => BoldPath,
            _ => RegularPath
        };

        if (!File.Exists(path))
            throw new FileNotFoundException($"폰트 파일을 찾을 수 없습니다: {path}");

        return File.ReadAllBytes(path);
    }

    // faceName: Family + 스타일 토큰
    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (!string.Equals(familyName, FamilyName, StringComparison.OrdinalIgnoreCase))
            return null; // 내가 처리할 패밀리만

        // 맑은 고딕은 italic 별도 파일이 없으니 regular/bold만 매핑
        return isBold
            ? new FontResolverInfo(FamilyName + "#b")
            : new FontResolverInfo(FamilyName + "#r");
    }
}
