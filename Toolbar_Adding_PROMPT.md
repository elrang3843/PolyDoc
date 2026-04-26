# IWPF 툴바 아이콘 세트 — Claude Code 작업 가이드

## 이 폴더 구조
```
icons_project/
├── PROMPT.md              ← 이 파일 (Claude Code 지시서)
├── ToolbarIcons.xaml      ← WPF ResourceDictionary (주요 아이콘 DrawingImage 포함)
└── icons/                 ← 개별 SVG 파일 80개 (종료 제외)
    ├── file_new.svg
    ├── file_open.svg
    ├── file_save.svg
    ├── file_saveas.svg
    ├── file_preview.svg
    ├── file_print.svg
    ├── file_close.svg
    ├── edit_copy.svg
    ├── edit_cut.svg
    ├── edit_delete.svg
    ├── edit_paste.svg
    ├── edit_import.svg
    ├── edit_export.svg
    ├── edit_docinfo.svg
    ├── edit_find.svg
    ├── edit_replace.svg
    ├── input_textbox_rect.svg
    ├── input_textbox_speech.svg
    ├── input_textbox_cloud.svg
    ├── input_textbox_spiky.svg
    ├── input_textbox_lightning.svg
    ├── input_table.svg
    ├── input_graph_line.svg
    ├── input_graph_pie.svg
    ├── input_graph_bar.svg
    ├── input_special_char.svg
    ├── input_formula.svg
    ├── input_emoji.svg
    ├── input_image.svg
    ├── input_sign.svg
    ├── shape_line.svg
    ├── shape_polyline.svg
    ├── shape_spline.svg
    ├── shape_square.svg
    ├── shape_rect.svg
    ├── shape_trapezoid.svg
    ├── shape_parallelogram.svg
    ├── shape_triangle_equil.svg
    ├── shape_triangle_isosc.svg
    ├── shape_triangle_right.svg
    ├── shape_triangle.svg
    ├── shape_circle.svg
    ├── shape_ellipse.svg
    ├── shape_arc.svg
    ├── shape_polygon.svg
    ├── shape_spline_fill.svg
    ├── shape_arrow.svg
    ├── format_font.svg
    ├── format_size.svg
    ├── format_bold.svg
    ├── format_italic.svg
    ├── format_underline.svg
    ├── format_strikethrough.svg
    ├── format_overline.svg
    ├── format_superscript.svg
    ├── format_subscript.svg
    ├── format_border.svg
    ├── format_text_color.svg
    ├── format_bg_color.svg
    ├── format_border_color.svg
    ├── format_char_width.svg
    ├── format_char_spacing.svg
    ├── format_line_spacing.svg
    ├── format_para_spacing.svg
    ├── format_indent.svg
    ├── format_outdent.svg
    ├── format_autonumber.svg
    ├── format_page.svg
    ├── format_color.svg
    ├── format_margin.svg
    ├── format_header.svg
    ├── format_footer.svg
    ├── format_columns.svg
    ├── tool_settings.svg
    ├── tool_dictionary.svg
    ├── tool_spellcheck.svg
    ├── tool_sign_create.svg
    ├── help_usage.svg
    ├── help_license.svg
    └── help_about.svg
```

---

## Claude Code에게 전달할 작업 지시

아래 내용을 Claude Code 터미널/채팅에 그대로 붙여넣으세요:

---

```
다음 작업을 수행해줘:

1. 이 폴더(icons_project/)를 C# 프로젝트의 Resources/Icons/ 경로로 복사해줘.
   - icons/ 폴더 전체를 Resources/Icons/svg/ 로
   - ToolbarIcons.xaml 을 Resources/ 로

2. .csproj 파일에 SVG 파일들을 리소스로 포함시켜줘:
   <ItemGroup>
     <Resource Include="Resources\Icons\svg\**\*.svg"/>
   </ItemGroup>

3. App.xaml의 Application.Resources에 ToolbarIcons.xaml 머지해줘:
   <ResourceDictionary.MergedDictionaries>
       <ResourceDictionary Source="Resources/ToolbarIcons.xaml"/>
   </ResourceDictionary.MergedDictionaries>

4. NuGet 패키지 Svg.Skia (또는 SharpVectors.Rendersvg) 를 추가해줘.
   SVG를 WPF Image로 로드하는 헬퍼 클래스 SvgIconHelper.cs 를 만들어줘:
   
   public static class SvgIconHelper
   {
       // SVG 파일을 ImageSource로 변환 (색상 currentColor → 지정색 치환 포함)
       public static ImageSource Load(string svgResourceKey, Color? color = null);
   }

5. 메인 툴바(MainToolbar.xaml 또는 MainWindow.xaml)에
   파일 메뉴 툴바 버튼 7개를 추가해줘:
   새파일/불러오기/저장/다른이름저장/미리보기/인쇄/닫기
   아이콘 크기는 20x20, 툴팁 텍스트 포함.

아이콘 파일명 → 한국어 툴팁 매핑:
file_new          → 새 파일 (IWPF)
file_open         → 불러오기
file_save         → 저장
file_saveas       → 다른 이름으로 저장
file_preview      → 미리보기
file_print        → 인쇄
file_close        → 닫기
edit_copy         → 복사
edit_cut          → 잘라내기
edit_delete       → 지우기
edit_paste        → 붙여넣기
edit_import       → 가져오기
edit_export       → 내보내기
edit_docinfo      → 문서 정보
edit_find         → 찾기
edit_replace      → 바꾸기
input_textbox_rect      → 글상자 (사각형)
input_textbox_speech    → 글상자 (말풍선)
input_textbox_cloud     → 글상자 (구름풍선)
input_textbox_spiky     → 글상자 (가시풍선)
input_textbox_lightning → 글상자 (번개상자)
input_table             → 표 (시트)
input_graph_line        → 꺾은선 그래프
input_graph_pie         → 파이 그래프
input_graph_bar         → 막대 그래프
input_special_char      → 특수문자
input_formula           → 수식
input_emoji             → 이모지
input_image             → 그림
input_sign              → 사인 (불러오기)
shape_line              → 직선
shape_polyline          → 폴리곤선
shape_spline            → 스프라인선
shape_square            → 정사각형
shape_rect              → 직사각형
shape_trapezoid         → 사다리꼴
shape_parallelogram     → 평행사변형
shape_triangle_equil    → 정삼각형
shape_triangle_isosc    → 이등변삼각형
shape_triangle_right    → 직각삼각형
shape_triangle          → 삼각형
shape_circle            → 원
shape_ellipse           → 타원
shape_arc               → 호
shape_polygon           → 폴리곤면
shape_spline_fill       → 스프라인면
shape_arrow             → 화살표
format_font             → 폰트
format_size             → 크기
format_bold             → 두껍게
format_italic           → 이탤릭
format_underline        → 밑줄
format_strikethrough    → 중간줄
format_overline         → 윗줄
format_superscript      → 위첨자
format_subscript        → 아래첨자
format_border           → 테두리
format_text_color       → 글자색
format_bg_color         → 배경색
format_border_color     → 줄 색
format_char_width       → 글자폭
format_char_spacing     → 글자 간격
format_line_spacing     → 줄 간격
format_para_spacing     → 문단 간격
format_indent           → 들여쓰기
format_outdent          → 내어쓰기
format_autonumber       → 자동번호
format_page             → 편집용지
format_color            → 색상
format_margin           → 여백
format_header           → 머릿글
format_footer           → 꼬릿글
format_columns          → 다단
tool_settings           → 설정
tool_dictionary         → 사전
tool_spellcheck         → 맞춤법 검사
tool_sign_create        → 사인 만들기
help_usage              → 사용 방법
help_license            → 라이센스
help_about              → About
```

---

## SVG 아이콘 설계 원칙

- **viewBox**: 모두 `0 0 24 24` 기준
- **색상**: `currentColor` 사용 → WPF의 Foreground/색상 테마 자동 반영
- **stroke**: 1.5px, round cap/join
- **강조색** (고정값):
  - 파란 액센트: `#378ADD`
  - 빨간 액센트: `#E24B4A`
  - 초록 액센트: `#1D9E75`
  - 앰버 액센트: `#BA7517`
- **크기**: 툴바 사용 시 `Width="20" Height="20"` 권장

## Git 커밋 제안

```bash
git add Resources/Icons/ Resources/ToolbarIcons.xaml
git commit -m "feat: add toolbar icon set (80 icons, SVG 24x24)"
git push
```
