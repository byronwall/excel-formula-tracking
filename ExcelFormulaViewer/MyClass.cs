/*
 * Created by SharpDevelop.
 * User: bwall
 * Date: 2/6/2017
 * Time: 4:16 PM
 * 
 * To change this template use Tools | Options | Coding | Edit Standard Headers.
 */
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Irony.Parsing;
using MathNet.Symbolics;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelFormulaViewer
{
	/// <summary>
	/// Description of MyClass.
	/// </summary>
	public class MyClass
	{
		
	}

	public static class RangeTools
	{
		static XLParser.FormulaAnalyzer GetParser(string formula)
		{
			if (formula[0] == '=') {
				formula = formula.Substring(1);
			}
			var parser = new XLParser.FormulaAnalyzer(formula);
			return parser;
		}

		static XLParser.FormulaAnalyzer GetParser(Excel.Range rng)
		{
			string formula = (string)rng.Formula;
			return GetParser(formula);
		}

		static Excel.Application app;
		static Excel.Worksheet ws;
		
		static string useSympyToCleanUpFormula(string expression)
		{
			ProcessStartInfo start = new ProcessStartInfo();
			start.FileName = @"C:\Users\TDAUser\AppData\Local\Continuum\Miniconda3\python.exe";
			
			string cmd = @"""C:/Documents/TDA Programming/UncertaintyCalcs/test_calcs.py""";
			
			start.Arguments = string.Format("{0} {1}", cmd, "-exp " + expression);
			start.UseShellExecute = false;
			start.RedirectStandardOutput = true;
			start.CreateNoWindow = true;
			
			using (Process process = Process.Start(start)) {
				using (StreamReader reader = process.StandardOutput) {
					string result = reader.ReadToEnd();
					Debug.Write(result);
					return result.Trim();
				}
			}
		}

		static void GenerateFullTreeWithReferences(ParseTreeNode root)
		{
			Stack<Tuple<ParseTreeNode, int>> nodes = new Stack<Tuple<ParseTreeNode, int>>();
			nodes.Push(Tuple.Create(root, 0));
			while (nodes.Count > 0) {
				var node = nodes.Pop();
				Debug.Print(new String(' ', node.Item2) + node.Item1);
				node.Item1.ChildNodes.ForEach(c => nodes.Push(Tuple.Create(c, node.Item2 + 1)));
				var astNode = node.Item1;
				if (astNode.Term.Name == "CellToken") {
					string cellRef = astNode.Token.ValueString;
					Excel.Range rngRef = (Excel.Range)ws.Range[cellRef];
					var subParser = GetParser(rngRef);
					nodes.Push(Tuple.Create(subParser.Root, node.Item2 + 1));
				}
			}
		}
		[ExcelCommand(MenuName = "Formula Tools", MenuText = "Parse Formulas")]
		public static void ParseFormula()
		{
			var _app = ExcelDnaUtil.Application;			
			app = (Excel.Application)_app;
			
			ws = (Excel.Worksheet)app.ActiveSheet;
			
			Excel.Range rng = (Excel.Range)app.ActiveCell;
			
			Debug.Print(rng.Formula.ToString());
			
			var parser = GetParser(rng);			
			var root = parser.Root;
			
			var newFormula = GetFormulaForFunc(root, true, true, -1);
			Debug.Print("processed formula: " + newFormula);
			
			MessageBox.Show(newFormula);
			
			var newParser = new XLParser.FormulaAnalyzer(newFormula);
			var noSumVersion = GetFormulaWithoutSum(newParser.Root);
			
			Debug.Print("no sum version: " + noSumVersion);
			
			useSympyToCleanUpFormula(noSumVersion);
			
			//take that formula and process as math
			
			var exp = Infix.ParseOrThrow(noSumVersion);
			
			Debug.Print("mathdotnet verison: " + Infix.Format(exp));
			Debug.Print("expanded verison: " + Infix.Format(Algebraic.Expand(exp)));
			Debug.Print("variables: " + string.Join(",", Structure.CollectIdentifierSymbols(exp).Select(c => c.Item)));
			
			//GenerateFullTreeWithReferences(root);			
		}
		
		[ExcelFunction(IsMacroType = true)]
		public static string GetFullFormulaOptions([ExcelArgument(AllowReference = true)]object arg, bool replaceRef = false, bool resolveName = false, int decimalPlaces = 5)
		{
			ExcelReference theRef = (ExcelReference)arg;
			Excel.Range rng = ReferenceToRange(theRef);
			
			ws = rng.Parent as Excel.Worksheet;
			
			var parser = GetParser(rng);			
			var root = parser.Root;
			
			var newFormula = GetFormulaForFunc(root, replaceRef, resolveName, -1);

			//remove the SUMs
			var noSumVersion = GetFormulaWithoutSum(new XLParser.FormulaAnalyzer(newFormula).Root);			
			var cleanFormula = Infix.Format(Infix.ParseOrThrow(noSumVersion));
			
			var finalFormula = cleanFormula;
			
			if (decimalPlaces > -1) {
				Debug.Print("Going back through a 2nd time");
				parser = GetParser(cleanFormula);
				var secondParserResult = GetFormulaForFunc(parser.Root, replaceRef, resolveName, decimalPlaces);
				finalFormula = Infix.Format(Infix.ParseOrThrow(secondParserResult));
			}
			
			//see if a short version of the formula is available
			var algFormula = Infix.Format(Algebraic.Expand(Infix.ParseOrThrow(finalFormula)));
			var ratFormula = Infix.Format(Rational.Expand(Infix.ParseOrThrow(finalFormula)));

			var shortFormula = new []{algFormula, ratFormula, finalFormula}.OrderBy(c=>c.Length).First();
	
			return shortFormula;	
		}
		
		[ExcelFunction(IsMacroType = true)]
		public static string GetFullFormula([ExcelArgument(AllowReference = true)]object arg)
		{			
			return GetFullFormulaOptions(arg, true, true);
		}
		
		static Excel.Range ReferenceToRange(ExcelReference xlref)
		{
			string refText = (string)XlCall.Excel(XlCall.xlfReftext, xlref, true);
			dynamic app = ExcelDnaUtil.Application;
			return app.Range[refText];
		}
		
		public static string GetFormulaForFunc(ParseTreeNode node, bool shouldReplaceRefWithConstant, bool shouldResolveNamedRange, int decimalPlaces)
		{
			var p1 = shouldReplaceRefWithConstant;
			var p2 = shouldResolveNamedRange;
			var p3 = decimalPlaces;
			
			//assume that the node is a "FunctionCall"
			try {
				
				
				switch (node.Term.Name) {
					case "FunctionCall":
						if (node.ChildNodes[0].Term.Name == "FunctionName") {
							//there will be arguments from here					
							var functionName = node.ChildNodes[0].ChildNodes[0].Token.ValueString; // this includes an opening (
							var functionArgs = GetFormulaForFunc(node.ChildNodes[1], p1, p2, p3);
							
							return string.Format("{0}{1})", functionName, functionArgs);
						} else {
							//there will be two items and a key in the middle

							if (node.ChildNodes[0].Term.Name == "Formula") {
							
								var first_term = GetFormulaForFunc(node.ChildNodes[0], p1, p2, p3);
								var oper = node.ChildNodes[1].Term.Name;
								var second_term = GetFormulaForFunc(node.ChildNodes[2], p1, p2, p3);
							
								return string.Format("({0}){1}({2})", first_term, oper, second_term);
							} else {
								var oper = node.ChildNodes[0].Term.Name;
								var second_term = GetFormulaForFunc(node.ChildNodes[1], p1, p2, p3);
							
								return string.Format("{0}({1})", oper, second_term);
							}
						}
					case "Arguments":
						return string.Join(",", node.ChildNodes.Select(c => GetFormulaForFunc(c, p1, p2, p3)));
					case "NumberToken":
						var numberValue = node.Token.ValueString;
						if (decimalPlaces > -1) {
							string formatString = String.Concat("0.", new string('#', decimalPlaces));
							Debug.Print(formatString);
							Debug.Print(numberValue);							
								
							var floatPretty = float.Parse(numberValue).ToString(formatString);
							
							Debug.Print(floatPretty);
							return floatPretty;
						}
						return numberValue;
					case "NameToken":
					case "NamedRangeCombinationToken":
						if (shouldResolveNamedRange) {
							//TODO this needs to actually resolve the name using Workbook.Names()
							return node.Token.ValueString;
						}
						return node.Token.ValueString;
					case "CellToken":
						//this will do an iterative ref, building a big formula as it goes						
						var cellToken = node.Token.ValueString;						
						Excel.Range rng = (Excel.Range)ws.Range[cellToken];
						if ((bool)rng.HasFormula) {
							return GetFormulaForFunc(GetParser(rng).Root, p1, p2, p3);
						}
						if (shouldReplaceRefWithConstant) {
							return rng.Value.ToString();
						}
						
						return cellToken;
					
					case "ReferenceFunctionCall":
						//this will do an iterative ref, building a big formula as it goes
						
						//first child is the start, second in the ":", third is the end cell
						
						//return a comma delim for each item in the range
						var firstCellAddr = node.ChildNodes[0].ChildNodes[0].ChildNodes[0].Token.ValueString;
						var secondCellAddr = node.ChildNodes[2].ChildNodes[0].ChildNodes[0].Token.ValueString;
						
						//get each cell in the range

						Excel.Range rngCells = (Excel.Range)ws.Range[firstCellAddr + ":" + secondCellAddr];						
						
						List<string> addresses = new List<string>();
						
						//this will replace the ranges with their cells 
						foreach (object rngCell in rngCells) {
							var rngCellType = rngCell as Excel.Range;
							Debug.Print(rngCellType.Address);
							
							if ((bool)rngCellType.HasFormula) {
								addresses.Add(GetFormulaForFunc(GetParser(rngCellType).Root, p1, p2, p3));
							} else {							
								addresses.Add(rngCellType.Address.Replace("$", ""));
							}
						}
						
						return string.Join(",", addresses.ToArray());
						
					default:
						//this handles all of the single node contains single node... it just goes down one level
						if (node.ChildNodes.Count > 0) {
							return GetFormulaForFunc(node.ChildNodes[0], p1, p2, p3);
						}
						return "";					
				}
			} catch (Exception e) {
				Debug.Print("error" + node);
				Debug.Print(e.ToString());
				return "";
			}			
		}
		
		public static string GetFormulaWithoutSum(ParseTreeNode node, string delim = ",")
		{
			try {
				switch (node.Term.Name) {
					case "FunctionCall":
						if (node.ChildNodes[0].Term.Name == "FunctionName") {
							//there will be arguments from here					
							var funcName = node.ChildNodes[0].ChildNodes[0].Token.ValueString;
							if (funcName == "SUM(") {
								return "" + GetFormulaWithoutSum(node.ChildNodes[1], "+");
							} else {
								return "" + funcName + GetFormulaWithoutSum(node.ChildNodes[1]) + ")";
							}
						} else {
							//there will be two items and a key in the middle					
							return "(" + GetFormulaWithoutSum(node.ChildNodes[0]) + ")" + node.ChildNodes[1].Term.Name + "(" + GetFormulaWithoutSum(node.ChildNodes[2]) + ")";
						}
					case "Arguments":
						return string.Join(delim, node.ChildNodes.Select(c => GetFormulaWithoutSum(c)));
					case "NumberToken":
						return node.Token.ValueString;
					case "CellToken":
						//this will do an iterative ref, building a big formula as it goes
						return node.Token.ValueString;
					case "NameToken":
					case "NamedRangeCombinationToken":
						//this will do an iterative ref, building a big formula as it goes
						return node.Token.ValueString;
						
					default:
						//this handles all of the single node contains single node... it just goes down one level
						if (node.ChildNodes.Count > 0) {
							return GetFormulaWithoutSum(node.ChildNodes[0]);			
						}
						return "";					
				}
			} catch (Exception e) {
				Debug.Print("error" + node);
				Debug.Print(e.ToString());
				return "";
			}			
		}
	}
}