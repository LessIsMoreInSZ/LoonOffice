using CommunityToolkit.Mvvm.ComponentModel;
using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Core;
using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LoonOffice
{
    public class MainWindowViewModel: ObservableObject
    {
        static void ProcessImages(Document doc, Application wordApp)
        {
            float pageWidth = doc.PageSetup.PageWidth - doc.PageSetup.LeftMargin - doc.PageSetup.RightMargin;
            float pageHeight = doc.PageSetup.PageHeight - doc.PageSetup.TopMargin - doc.PageSetup.BottomMargin;
            float verticalThreshold = pageHeight * 0.2f; // 剩余空间阈值（页面高度的20%）

            // 遍历所有内联图片
            foreach (InlineShape inlineShape in doc.InlineShapes)
            {
                if (inlineShape.Type == WdInlineShapeType.wdInlineShapePicture)
                {
                    // 转换为浮动图形以便精确控制位置
                    inlineShape.ConvertToShape();
                }
            }

            // 处理所有浮动图形
            foreach (Shape shape in doc.Shapes)
            {
                try
                {
                    if (shape.Type == MsoShapeType.msoPicture ||
                        shape.Type == MsoShapeType.msoLinkedPicture)
                    {
                        // 获取图片锚点所在的页面
                        int pageNumber = shape.Anchor.Information[WdInformation.wdActiveEndPageNumber];

                        // 获取当前页剩余空间
                        float verticalPosition = shape.Anchor.Information[WdInformation.wdVerticalPositionRelativeToPage];
                        float remainingSpace = pageHeight - verticalPosition;

                        // 检查是否需要换页
                        if (remainingSpace < verticalThreshold)
                        {
                            InsertPageBreak(shape.Anchor);
                            // 更新锚点到新页面
                            shape.Anchor = shape.Anchor.Paragraphs[1].Range;
                            pageNumber = shape.Anchor.Information[WdInformation.wdActiveEndPageNumber];
                        }

                        // 设置图片属性
                        shape.LockAspectRatio = MsoTriState.msoTrue; // 保持纵横比
                        shape.Width = pageWidth; // 适应页面宽度

                        // 计算垂直位置（居中）
                        float topPosition = (pageHeight - shape.Height) / 2 + doc.PageSetup.TopMargin;

                        // 设置位置（相对于页面）
                        shape.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage;
                        shape.Top = topPosition;

                        // 设置水平居中
                        shape.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage;
                        shape.Left = (doc.PageSetup.PageWidth - shape.Width) / 2;

                        // 设置文字环绕方式
                        shape.WrapFormat.Type = WdWrapType.wdWrapTopBottom;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"处理图片时出错: {ex.Message}");
                }
            }
        }

        static void InsertPageBreak(Microsoft.Office.Interop.Word.Range range)
        {
            // 在指定范围前插入分页符
            range.InsertBreak(Type: WdBreakType.wdPageBreak);

            // 将锚点移动到新页面
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
            range.InsertParagraphAfter();
            range.Collapse(WdCollapseDirection.wdCollapseEnd);
        }
    }
}
