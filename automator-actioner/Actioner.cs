using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json.Nodes;
using System.Threading.Tasks;

using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.Loader;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using automator_baselib;


namespace automator_actioner
{
    public class Actioner
    {

        string sourceFileName = "";
        string destDirectory = "";
        // 設定値
        public Dictionary<string, string> properties { set; get; } = new Dictionary<string, string>();

        // プラグインdll
        Dictionary<string, PluginBase> dlls = new Dictionary<string, PluginBase>();

        Dictionary<string, List<string>> programs = new Dictionary<string, List<string>>();
        string casePtr = "";        // 実行中のケース名
        public int ptr = 0;                // ケース中のプログラム番号

        public Actioner() { }

        public IEnumerable<string> Init()
        {
            // プラグイン読み込み
            string libPath = Path.Combine(Directory.GetCurrentDirectory(), "libs");

            foreach (string lib in Directory.GetFiles(libPath))
            {
                Assembly assembly = Assembly.LoadFrom(lib);
                foreach (Type type in assembly.GetTypes())
                {
                    if (type == null) { continue; }
                    // クラス名にAutomatorLibraryがあるもののみインスタンス化
                    if (type.FullName.EndsWith("AutomatorLibrary"))
                    {
                        yield return new Output("D-cmn-0014", "ライブラリ[" + type.Namespace + "]の読み込みを開始します。", "", "").toString();

                        // インスタンス化
                        Type libType = assembly.GetType(type.Namespace + ".AutomatorLibrary");
                        PluginBase plugin = (PluginBase)Activator.CreateInstance(libType);

                        foreach (string log in plugin.Init())
                        {
                            String er = checkLog(log, "初期化処理");
                            yield return log;
                            if (er != "")
                            {
                                yield return log;
                                yield return new Output("E-cmn-0020", "ライブラリ[" + type.Namespace + "]の読み込みに失敗しました。", "", "").toString();
                                yield break;
                            }
                        }

                        dlls.Add(plugin.identifier, plugin);

                        yield return new Output("I-cmn-0015", "ライブラリ[" + type.Namespace + "]が正常に読み込まれました。", "", "").toString();
                    }
                }
            }

            // 共通設定ファイル読み込み
            string err = ReadConfigFile(Path.Combine(Directory.GetCurrentDirectory(), "config", "common.ini"));
            if (err == "")
            {
                yield return new Output("I-cmn-0019", "設定値を読み込みました。", "", "").toString();
            }
            else 
            {
                yield return err;
            }

            // tmpフォルダ作成
            string tmpDir = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "tmp");
            if (!Directory.Exists(tmpDir))
            {
                Directory.CreateDirectory(tmpDir);
            }
        }

        public void SetSourceFile(string fileName)
        {
            this.sourceFileName = fileName;
        }

        public void SetDestDirectory(string folderName)
        {
            this.destDirectory = folderName;
        }

        public string GetDestDirectory() { return this.destDirectory; }

        public string ParseProgram()
        {
            string ext = System.IO.Path.GetExtension(this.sourceFileName);
            List<string> program = new List<string>();


            if (ext == ".xlsx")
            {
                // excelの場合
                return new Output("E-cmn-0000", "Excel形式は未実装！", "", "").toString();

            }
            else if (ext == ".csv")
            {
                // csvの場合
                return ReadCsvFile();
            }
            else
            {
                // unsupported protocol
                return new Output("E-cmn-0014", "入力ファイルがサポートされていない形式です。 [EXT: " + ext + "]", "", "").toString();
            }
        }

        public IEnumerable<string> Finish()
        {
            // 各dllのリセットを実行
            foreach (var dll in this.dlls)
            {
                foreach (string log in dll.Value.Reset())
                {
                    yield return log;
                }
                yield return new Output("D-cmn-0028", "dll[" + dll.Key + "]をリセット", "", "").toString();
            }

            // ディレクトリ作成，tmpから移動
            if (this.destDirectory != null && this.destDirectory != "")
            {
                yield return new Output("D-cmn-0039", "ディレクトリ移動処理を開始します。", "", "").toString();
                if (Directory.Exists(this.destDirectory))
                {
                    string tmpDir = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "tmp");
                    string outputDir = System.IO.Path.Combine(this.destDirectory, this.casePtr);

                    string err = "";
                    try
                    {
                        try
                        {
                            Directory.Delete(outputDir, true);
                        } catch
                        {
                        }

                        Directory.Move(tmpDir, outputDir);

                        Directory.CreateDirectory(tmpDir);
                    }
                    catch (Exception ex)
                    {
                        err = new Output("E-cmn-0030", "ファイル移動中にエラーが発生しました。" + ex.Message, "", "").toString();
                    }
                    if (err != "")
                    {
                        yield return err;
                    }
                }
                else
                {
                    yield return new Output("E-cmn-0029", "結果出力先が存在しませんでした。", "", "").toString();
                    yield break;
                }
            }
        }

        public void clearProgram()
        {

        }

        public string[] GetAllTestCaseNames()
        {
            foreach (string a in this.programs.Keys.ToArray()) { Debug.WriteLine("A: " + a); }
            Debug.WriteLine(this.programs.Keys.ToArray());
            return this.programs.Keys.ToArray();
        }

        private string ReadCsvFile()
        {
            this.programs = new Dictionary<string, List<string>>();

            using (StreamReader reader = new StreamReader(this.sourceFileName))
            {
                string testCaseName = "";

                while (!reader.EndOfStream)
                {
                    List<string> program = new List<string>();
                    string line = reader.ReadLine();
                    string[] values = line.Split(',');

                    // バリデーションエラー
                    if (values.Length < 1) { return new Output("E-cmn-0009", "入力csv内に空行があります。", "", "").toString(); }
                    testCaseName = values[0];
                    if (testCaseName == "") { return new Output("E-cmn-0010", "入力csv内に空のケース名があります。", "", "").toString(); }
                    if (this.programs.ContainsKey(values[0])) { return new Output("E-cmn-0011", "入力csv内で重複したケース名があります。[ケース名: " + values[0] + "]", "", "").toString(); }
                    if (values.Length < 2) { return new Output("E-cmn-0012", "プログラムがありません。", "", "").toString(); }

                    // 読み込み
                    for (int i = 1; i < values.Length; i++)
                    {
                        // #から始まる行はコメント
                        if (!values[i].StartsWith("#"))
                        {
                            program.Add(values[i].Replace(" ", ""));
                        }
                    }

                    programs.Add(testCaseName, program);
                }

            }
            return "";
        }

        public void setPtr(int ptr) { this.ptr = ptr; }

        public void setCasePtr(string casePtr) { this.casePtr = casePtr; }

        public string getActionCmd()
        {
            if (this.ptr >= this.programs[this.casePtr].Count)
            {
                return "__EOA__";
            }
            else
            {
                return this.programs[this.casePtr][this.ptr];
            }
        }

        public void NextPtr()
        {
            if (this.ptr < this.programs[this.casePtr].Count)
            {
                this.ptr++;
            }
        }

        public IEnumerable<string> Action(string cmd)
        {
            string[] cmds = cmd.Split("->");
            string input = "";

            foreach (string c in cmds)
            {
                // コマンドが5文字以内ならエラー
                if (c.Length < 5)
                {
                    yield return new Output("E-cmn-0016", "コマンドが短すぎます。", "", "").toString();
                    yield break;
                }

                string lib = c.Substring(0, 3);

                // libが見つからなければエラー
                if (!dlls.ContainsKey(lib))
                {
                    yield return new Output("E-cmn-0017", "該当のライブラリが見つかりませんでした。", "", "").toString();
                    yield break;
                }

                // 実行
                foreach (string result in this.dlls[lib].Action(input, c.Substring(5)))
                {
                    Debug.WriteLine("RESULT: " + result); 
                    yield return result;

                    string s = "";
                    try
                    {
                        // outputを次のinputにまわす
                        JsonNode jsonNode = JsonNode.Parse(result);
                        string outputType = jsonNode["outputValue"]["type"]?.ToString();
                        string outputText = jsonNode["outputValue"]["content"]?.ToString();
                        input = "{\"type\":\"" + outputType + "\",\"content\":\"" + outputText + "\"}";
                    }
                    catch (Exception e)
                    {
                        s = new Output("E-cmn-0018", "output定義異常です。", "", "").toString();
                    }

                    if (s != "")
                    {
                        yield return s;
                        yield break;
                    }
                }
            }
        }

        private string ReadConfigFile(string configPath)
        {
            if (!File.Exists(configPath))
            {
                return new Output("F-cmn-9003", "設定ファイルが見つかりません。[configPath:  " + configPath + "]", "", "").toString();
            }

            foreach (string line in File.ReadAllLines(configPath))
            {
                if (string.IsNullOrWhiteSpace(line) || line.TrimStart().StartsWith("#")) continue;
                string[] d = line.Split(new char[] { '=' });
                if (d.Length == 2)
                {
                    this.properties[d[0].Trim()] = d[1].Trim();
                }
                else
                {
                    return new Output("F-cmn-9004", "設定ファイルに異常な値が見つかりました。[configFilePath: " + configPath + " line:  " + line + "]", "", "").toString();
                }
            }
            return "";
        }

        private string checkLog(string log, string processName)
        {

            // 結果のjsonパース
            JsonNode? jsonNode;
            try
            {
                jsonNode = JsonNode.Parse(log);
                if (jsonNode == null)
                {
                    return new Output("E-cmn-0004", "結果の取得に失敗しました。 {log: null}", "", "").toString();
                }
            }
            catch
            {
                return new Output("E-cmn-0005", "結果の取得に失敗しました。 {log: " + log + "}", "", "").toString();
            }

            string errCode = "";
            string errMsg = "";
            try
            {
                errCode = jsonNode["errCode"]?.ToString();
                if (errCode == "")
                {
                    return new Output("E-cmn-0008", "errCodeの取得に失敗しました。", "", "").toString();
                }
            }
            catch
            {
                return new Output("E-cmn-0006", "errCodeの取得に失敗しました。", "", "").toString();
            }
            try
            {
                errMsg = jsonNode["errMsg"]?.ToString();
            }
            catch
            {
                return new Output("E-cmn-0007", "errMsgの取得に失敗しました。", "", "").toString();
            }

            if (errCode[0] == 'E')
            {
                return new Output("E-cmn-0007", "エラーのため" + processName + "を中断します。", "", "").toString();
            }

            return "";

        }
    }

}
