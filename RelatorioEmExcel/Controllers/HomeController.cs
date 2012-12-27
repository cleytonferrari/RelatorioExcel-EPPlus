using System;
using System.Data;
using System.Drawing;
using System.IO;
using System.Web.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace RelatorioEmExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Excel()
        {
            var tbl = CriaDataTableProdutos();

            IncluirNoDataTable("Arroz Tipo 1", "PCT", 12, 2, tbl);
            IncluirNoDataTable("Feijão São João", "PCT", 7, 3, tbl);
            IncluirNoDataTable("Macarrão Ferrari", "PCT", 8, (decimal) 2.3, tbl);


            using (var pacote = new ExcelPackage())
            {
                //Cria uma pasta de Trabalho, no arquivo do Excel
                var ws = pacote.Workbook.Worksheets.Add("Produtos");

                //Carrega os dados para a planilha, inicia na celula A1, a primeira linha é o descricao dos campos usado na tabela
                ws.Cells["A1"].LoadFromDataTable(tbl, true);

                //Formata a coluna para numerico, valor unitario (4)
                using (var col = ws.Cells[2, 4, 2 + tbl.Rows.Count, 4])
                {
                    col.Style.Numberformat.Format = "#,##0.00";
                    col.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                }

                //Formata o cabeçalho, 4 celulas (a1:d1)
                using (var rng = ws.Cells["A1:F1"])
                {
                    rng.Style.Font.Bold = true;
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                    rng.Style.Font.Color.SetColor(Color.White);
                }
                //Seta um texto para o Valor Total
                ws.Cells["F1"].Value = "Valor Total";

                //Formata a coluna do valor total, para numeric, e coloca uma formula
                for (var i = 2; i <= tbl.Rows.Count + 1; i++)
                {
                    ws.Cells["F" + i].Style.Numberformat.Format = "#,##0.00";
                    ws.Cells["F" + i].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
                    ws.Cells["F" + i].Formula = string.Format(" {0}*{1}", "D" + i, "E" + i);
                }

                //Formata o valor total
                var vTotal = tbl.Rows.Count + 2;
                ws.Cells["F" + vTotal].Style.Font.Bold = true;
                ws.Cells["F" + vTotal].Style.Fill.PatternType = ExcelFillStyle.Solid;
                ws.Cells["F" + vTotal].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189));
                ws.Cells["F" + vTotal].Style.Font.Color.SetColor(Color.White);
                ws.Cells["F" + vTotal].Style.Numberformat.Format = "#,##0.00";
                ws.Cells["F" + vTotal].Formula = string.Format("SUM({0},{1})", "F2", "F" + (vTotal - 1));


                //gera o arquivo para download
                var stream = new MemoryStream();
                pacote.SaveAs(stream);
                const string nomeDoArquivo = "RelatorioDeProdutos.xlsx";
                const string contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                stream.Position = 0;
                return File(stream, contentType, nomeDoArquivo);
            }

        }



        private static void IncluirNoDataTable(string descricao, string unidadeMedidada, int quantidade, decimal valorUnitario, DataTable mTable)
        {
            var linha = mTable.NewRow();

            linha["Código"] = Guid.NewGuid().ToString();
            linha["Descrição"] = descricao;
            linha["Unidade de Medida"] = unidadeMedidada;
            linha["Quantidade"] = quantidade;
            linha["Valor Unitário"] = valorUnitario;

            mTable.Rows.Add(linha);
        }

        private static DataTable CriaDataTableProdutos()
        {
            var mDataTable = new DataTable();

            var mDataColumn = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Código"
            };
            mDataTable.Columns.Add(mDataColumn);

            mDataColumn = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Descrição"
            };
            mDataTable.Columns.Add(mDataColumn);

            mDataColumn = new DataColumn
            {
                DataType = Type.GetType("System.String"),
                ColumnName = "Unidade de Medida"
            };
            mDataTable.Columns.Add(mDataColumn);

            mDataColumn = new DataColumn
            {
                DataType = Type.GetType("System.Int32"),
                ColumnName = "Quantidade"
            };
            mDataTable.Columns.Add(mDataColumn);

            mDataColumn = new DataColumn
            {
                DataType = Type.GetType("System.Decimal"),
                ColumnName = "Valor Unitário"
            };
            mDataTable.Columns.Add(mDataColumn);

            return mDataTable;
        }

    }
}
