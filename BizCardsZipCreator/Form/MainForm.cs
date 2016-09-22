using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.IO.Compression;
using jp.tabamotch.BizCardsZipCreator.Utility;

namespace jp.tabamotch.BizCardsZipCreator.Forms
{
    /// <summary>
    /// 
    /// </summary>
    public partial class MainForm : Form
    {
        private readonly Dictionary<string, string> _contactsDic;
        private readonly Dictionary<string, int> _imageNameColumnIndex;
        private readonly Dictionary<string, int> _imageCellColumnIndex;
        private string _tempFolderPath;

        private List<IDisposable> _disposeItems = new List<IDisposable>();

        /// <summary>
        /// コンストラクタ
        /// </summary>
        public MainForm()
        {
            InitializeComponent();

            // イメージ名の列番号を格納するDicの初期化
            _imageNameColumnIndex = new Dictionary<string, int>();

            // イメージサムネイルを表示する列のDic初期化
            _imageCellColumnIndex = new Dictionary<string, int>();

            // ヘッダ部分Dicの初期化
            _contactsDic = CreateContactsTextColumnDic();

            // DataGridViewの初期化
            InitializeGrid();
        }

        private void InitializeGrid()
        {
            // DataGridViewのヘッダの初期化(テキスト手入力部)
            foreach (KeyValuePair<string, string> k in _contactsDic)
            {
                int colNumber = dataGridView1.Columns.Add(k.Key, k.Value);

                if (k.Key.Contains("image"))
                {
                    dataGridView1.Columns[colNumber].Visible = false;
                    _imageNameColumnIndex.Add(k.Key, colNumber);
                }
            }

            // 名刺画像表示列 表
            DataGridViewImageColumn colImage1 = new DataGridViewImageColumn()
            {
                Name = "名刺画像表示(表)"
            };

            dataGridView1.Columns.Add(colImage1);

            // 名刺画像選択用カラム追加
            DataGridViewButtonColumn colButtonImage1 = new DataGridViewButtonColumn()
            {
                Name = "imageButton",
                HeaderText = "名刺選択(表)",
                UseColumnTextForButtonValue = true,
                Text = "イメージを選択",
                Tag = "image"
            };

            int colNumImageButton1 = dataGridView1.Columns.Add(colButtonImage1);

            // 名刺画像表示列 裏
            DataGridViewImageColumn colImage2 = new DataGridViewImageColumn()
            {
                Name = "名刺画像表示(裏)",

            };

            dataGridView1.Columns.Add(colImage2);

            DataGridViewButtonColumn colButtonImage2 = new DataGridViewButtonColumn()
            {
                Name = "image2Button",
                HeaderText = "名刺選択(裏)",
                UseColumnTextForButtonValue = true,
                Text = "イメージを選択",
                Tag = "image2"
            };

            int colNumImageButton2 = dataGridView1.Columns.Add(colButtonImage2);

            // イメージ選択ボタンを押下したときのイベントを設定
            dataGridView1.CellClick += (sender, e) =>
            {
                if (e.ColumnIndex == colNumImageButton1 ||
                    e.ColumnIndex == colNumImageButton2)
                {
                    CellImageButtonClicked(e, colNumImageButton1, colNumImageButton2);
                }
            };
        }

        /// <summary>
        /// 画像選択ボタン押下時イベント
        /// </summary>
        /// <param name="e">イベントパラメータ</param>
        /// <param name="colNumImageButton1">画像選択ボタン(表)の列インデックス</param>
        /// <param name="colNumImageButton2">画像選択ボタン(裏)の列インデックス</param>
        private void CellImageButtonClicked(DataGridViewCellEventArgs e, int colNumImageButton1, int colNumImageButton2)
        {
            int clickedColumnIndex = e.ColumnIndex;
            int clickedRowIndex = e.RowIndex;

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    Tuple<bool, string, string> t = CopyImageFileToTemp(dialog.FileName);
                    if (t.Item1)
                    {
                        string imageFileName = t.Item2;

                        string tag = (clickedColumnIndex == colNumImageButton1) ? "image" : "image2";
                        dataGridView1[_imageNameColumnIndex[tag], clickedRowIndex].Value = imageFileName;

                        // 動かないのでコメントアウト中
                        //((DataGridViewButtonCell)dataGridView1[clickedColumnIndex, clickedRowIndex]).UseColumnTextForButtonValue = true;
                        //((DataGridViewButtonCell)dataGridView1[clickedColumnIndex, clickedRowIndex]).Value = "選択済み";
                    }
                    else
                    {
                        MessageBox.Show(this, t.Item3, "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
            }
        }

        /// <summary>
        /// 画面ロードイベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_Load(object sender, EventArgs e)
        {
            // 一時フォルダパス
            _tempFolderPath = Path.GetRandomFileName();
        }

        /// <summary>
        /// 画面終了時イベント
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            // DisposeするものはDisposeする
            _disposeItems?.ForEach((item) => item.Dispose());

            if (Directory.Exists(_tempFolderPath))
            {
                try
                {
                    // 一時ディレクトリを削除
                    Directory.Delete(_tempFolderPath, true);
                }
                catch
                {
                    // 何もしない
                }
            }
        }

        /// <summary>
        /// テキスト入力列の組み合わせ
        /// CSVファイルのヘッダとDataGridViewのヘッダの名称を結びつける為のDictionaryを作成
        /// </summary>
        /// <returns>CSVヘッダとDataGridViewヘッダのDictionary</returns>
        private Dictionary<string, string> CreateContactsTextColumnDic()
        {
            Dictionary<string, string> result = new Dictionary<string, string>
            {
                {"Last", "姓"},
                {"First", "名"},
                {"LastP", "姓の読み"},
                {"FirstP", "名の読み"},
                {"Organization", "会社"},
                {"Department", "部署"},
                {"JobTitle", "役職"},
                {"Note", "メモ"},
                {"Email_work", "Eメール(勤務先)"},
                {"image", "名刺画像(表)"},
                {"image2", "名刺画像(裏)"},
            };

            return result;
        }

        // ファイル名に付与する連番
        private int _fileCount = 0;

        // イメージファイル名(固定)
        private const string IMAGE_FILE_DEFAULT_NAME = "image{0:00000}{1}";
         
        /// <summary>
        /// 画像を一時フォルダにコピー
        /// </summary>
        /// <param name="filePath">コピー元画像ファイルパス</param>
        /// <returns>処理結果タプル(成功/失敗、成功した場合のファイル名、失敗した場合のエラーメッセージ)</returns>
        private Tuple<bool, string, string> CopyImageFileToTemp(string filePath)
        {
            // 拡張子取得
            FileInfo file = new FileInfo(filePath);
            if(!file.Exists)
            {
                return new Tuple<bool, string, string>(false, null, "ファイルが見つかりません。");
            }

            try
            {
                using (Image image = Image.FromFile(filePath))
                {
                    int selectedRow = dataGridView1.SelectedCells[0].RowIndex;
                    int selectedColumn = dataGridView1.SelectedCells[0].ColumnIndex;

                    ((DataGridViewImageCell)dataGridView1[selectedColumn - 1, selectedRow]).ValueIsIcon = false;

                    // 縮小イメージを取得
                    Image convertedImage = ImageUtility.ConvertImageSize(image);

                    // 画面を閉じるときにDisposeできるようにしておく。
                    _disposeItems.Add(convertedImage);

                    // 縮小イメージを設定
                    dataGridView1[selectedColumn - 1, selectedRow].Value = convertedImage;
                }
            }
            catch (Exception)
            {
                return new Tuple<bool, string, string>(false, null, "選択されたファイルはイメージファイルではありません。");
            }

            // 新しいイメージファイル名生成
            string currentImageFileName = string.Format(IMAGE_FILE_DEFAULT_NAME, _fileCount++, file.Extension);

            if(!Directory.Exists(_tempFolderPath))
            {
                // 一時フォルダがない場合は作成
                Directory.CreateDirectory(_tempFolderPath);
            }

            try
            {
                // 一時フォルダにイメージファイルをコピー
                File.Copy(filePath, Path.Combine(_tempFolderPath, currentImageFileName));
            }
            catch(Exception)
            {
                return new Tuple<bool, string, string>(false, null, "イメージファイルのコピーに失敗しました。");
            }

            // 処理に成功した時は新しいイメージファイル名を返す
            return new Tuple<bool, string, string>(true, currentImageFileName, null);
        }

        /// <summary>
        /// 行追加ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addRowButton_Click(object sender, EventArgs e)
        {
            // 行追加
            dataGridView1.Rows.Add();

            // 行削除ボタンを押下可能にする
            deleteRowButton.Enabled = true;
        }

        /// <summary>
        /// 行削除ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void deleteRowButton_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show(this, "行を削除してもよろしいですか？", "確認", MessageBoxButtons.OKCancel, MessageBoxIcon.Question);

            if(res != DialogResult.OK)
            {
                // 確認ダイアログで「キャンセル」押下時、処理を抜ける
                return;
            }

            int selectedRowIndex = dataGridView1.SelectedCells[0].RowIndex;
            if(selectedRowIndex >= 0)
            {
                // 行削除
                dataGridView1.Rows.RemoveAt(selectedRowIndex);
            }

            // 行削除ボタンの押下可否制御。行が残っている場合は押下可能、行が0行なら押下不可
            deleteRowButton.Enabled = dataGridView1.Rows.Count > 0;
        }

        /// <summary>
        /// ZIPファイル作成ボタン押下処理
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void createZipButton_Click(object sender, EventArgs e)
        {
            // ZIPファイルを作成する為の一時フォルダ
            // メソッドを抜ける時に削除する
            string tempDir = Path.GetRandomFileName();

            try
            {
                // 一時フォルダ作成
                Directory.CreateDirectory(tempDir);

                // CSVファイルのファイル情報生成。ファイル名は決まっている為、固定
                FileInfo contactsCsv = new FileInfo(Path.Combine(tempDir, "contacts.csv"));

                using (StreamWriter writer = contactsCsv.CreateText())
                {
                    // CSVファイルのヘッダ
                    writer.WriteLine(string.Join(",", _contactsDic.Keys));

                    // カラムの列番号取得用
                    // 本当なら必要ないかもしれないけど、後々列を追加するときの為に入れておく
                    Dictionary<string, int> colIndexes = new Dictionary<string, int>();
                    foreach (string key in _contactsDic.Keys)
                    {
                        for (int j = 0; j < dataGridView1.Columns.Count; j++)
                        {
                            if (dataGridView1.Columns[j].Name == key)
                            {
                                colIndexes.Add(key, j);
                                break;
                            }
                        }
                    }

                    // DataGridViewの行数だけ繰り返し
                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        List<string> line = new List<string>();

                        foreach(string key in colIndexes.Keys)
                        {
                            // セルの入力値を取得
                            string value = (string)dataGridView1[colIndexes[key], i].Value;

                            if(string.IsNullOrWhiteSpace(value))
                            {
                                // 列に何も入力されていない場合は、空文字を設定
                                line.Add(string.Empty);
                            }
                            else
                            {
                                if (key.Contains("image"))
                                {
                                    // イメージファイルが紐づけられていれば、一時フォルダにコピー
                                    // ZIPファイルを再出力するかも
                                    File.Copy(Path.Combine(_tempFolderPath, value), Path.Combine(tempDir, value));
                                }

                                // 列に値が設定されている場合はその値を出力する
                                // ダブルクォーテーションを2つにする
                                //value = value.Replace("\"", "\"\"");
                                // 出力する文字列をダブルクォーテーションで囲む
                                //value = $"\"{value}\"";

                                // 2016/09/19
                                // CSVファイルの改行・カンマ・ダブルクォーテーション対応を行ったが、
                                // BizCards側が対応していなかった為、コメントアウト

                                // 出力用リストに追加
                                line.Add(value);
                            }
                        }

                        // ファイルに出力
                        writer.WriteLine(string.Join(",", line));
                    }

                    writer.Flush();
                    writer.Close();
                }

                using (SaveFileDialog dialog = new SaveFileDialog())
                {
                    string test = dialog.InitialDirectory;

                    // BizCardsで取り込めるファイル名は決まっているので、あらかじめ設定しておく
                    dialog.FileName = "contacts.zip";
                    dialog.Filter = "ZIPファイル(*.zip)|*.zip";

                    if (dialog.ShowDialog(this) == DialogResult.OK)
                    {
                        // ZIPファイル作成
                        ZipFile.CreateFromDirectory(tempDir, dialog.FileName, CompressionLevel.Fastest, false);

                        // メッセージを表示して終了
                        MessageBox.Show(this, "処理が完了しました。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            finally
            {
                try
                {
                    // 後始末として一時フォルダを削除
                    Directory.Delete(tempDir, true);
                }
                catch
                {
                    // 削除に失敗しても何もしない
                }
            }
        }
    }
}
