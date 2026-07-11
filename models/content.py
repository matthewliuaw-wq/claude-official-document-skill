from pydantic import BaseModel, Field
from typing import Optional, List


class CoverPageContent(BaseModel):
    title: str = Field(..., description="文档标题")
    subtitle: Optional[str] = Field(None, description="副标题")
    info_lines: Optional[List[str]] = Field(None, description="封面附加信息行（单位、日期等）")
    show_cover: bool = Field(True, description="是否生成封面")


class TableContent(BaseModel):
    caption: Optional[str] = Field(None, description="表格标题（显示在表格上方）")
    col_widths: Optional[List[int]] = Field(None, description="列宽（DXA 单位），不设则自动")
    row_heights: Optional[List[int]] = Field(None, description="各行最小高度（DXA 单位），不设则自动")
    headers: List[str] = Field(..., description="表头行文本")
    rows: List[List[str]] = Field(..., description="数据行，每行是字符串列表")


class BulletListContent(BaseModel):
    style: str = Field("bullet", description="列表样式：bullet 或 numbered")
    items: List[str] = Field(..., description="列表项文本")


class CodeBlockContent(BaseModel):
    language: Optional[str] = Field(None, description="编程语言，如 python、java")
    code: str = Field(..., description="代码内容")


class FigureContent(BaseModel):
    number: str = Field(..., description="图号，如 1-1、2.3")
    caption: str = Field(..., description="图名/图标题")
    image_path: Optional[str] = Field(None, description="图片文件路径")


class ContentBlock(BaseModel):
    type: str = Field(..., description="块类型：heading/paragraph/table/bullet_list/code_block/figure/page_break/toc/attachment_label")
    level: Optional[int] = Field(None, description="标题级别 1-3，type=heading 时必填", ge=1, le=3)
    text: Optional[str] = Field(None, description="heading/paragraph 的文本内容")
    bold_prefix: Optional[str] = Field(None, description="段落中加粗的前缀文字")
    table: Optional[TableContent] = Field(None, description="表格数据，type=table 时必填")
    bullet_list: Optional[BulletListContent] = Field(None, description="列表数据，type=bullet_list 时必填")
    code_block: Optional[CodeBlockContent] = Field(None, description="代码块数据，type=code_block 时必填")
    figure: Optional[FigureContent] = Field(None, description="图片数据，type=figure 时必填")
    alignment: Optional[str] = Field(None, description="段落对齐覆盖：left/center/right/justified")
    space_before: Optional[int] = Field(None, description="段前间距（Pt值），用于落款前的间距控制")


class DocumentMeta(BaseModel):
    title: str = Field(..., description="文档短标题（用于页眉页脚）")
    header_text: Optional[str] = Field(None, description="页眉文字，默认使用 title")


class DocumentContent(BaseModel):
    meta: DocumentMeta = Field(..., description="文档元信息")
    cover: Optional[CoverPageContent] = Field(None, description="封面内容，不设则跳过封面")
    content: List[ContentBlock] = Field(..., description="有序内容块列表")
