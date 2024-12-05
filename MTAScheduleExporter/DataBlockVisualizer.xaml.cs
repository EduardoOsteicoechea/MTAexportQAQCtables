using System.Windows;
using System.Windows.Documents;
using System.Windows.Media;

namespace MTAScheduleExporter
{
   public partial class DataBlockVisualizer: Window
   {
      public DataBlockVisualizer(string data)
      {
         InitializeComponent();

         FlowDocument richTextBoxFlowDocument = DataRichTextBox.Document;

         Paragraph paragraph = richTextBoxFlowDocument.Blocks.FirstBlock as Paragraph;
         paragraph.LineHeight = 10;
         paragraph.Margin = new Thickness(5,0,5,0);
         paragraph.TextAlignment = TextAlignment.Left;
         paragraph.FontSize = 14;
         paragraph.FontFamily = new FontFamily("Arial");
          
         DataRichTextBox.AppendText(data);
      }
   }
}