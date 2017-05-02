using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using Ionic.Zip;
using System.Reflection;
using System.Threading;

namespace UnZipper
{
    class Program
    {
        #region "定数"
        /// <summary>作業フォルダ</summary>
        static string _WORK_FOLDER = "";
        /// <summary>処理済フォルダ</summary>
        static string _PROCESSED_FOLDER = "";

        #endregion "定数"

        private static System.Threading.Timer _wTimer;


        static void Main(string[] args)
        {
            //作業フォルダ作成
            _WORK_FOLDER = Settings1.Default.作業フォルダ;
            _PROCESSED_FOLDER = Path.Combine(_WORK_FOLDER, "PROCESSED");
            CreateFolder(_WORK_FOLDER);
            CreateFolder(_PROCESSED_FOLDER);

            //引数が存在するか確認。引数０の場合ｅｘｅ実行であると判断
            var targets = GetArgs(args);
            var monitoring = targets.Count() == 0;

            //引数のzipたちを解凍
            UnZip(targets);
            if ( ! monitoring )
            {
                Console.WriteLine("処理完了");
                return;
            }

            //モニタ有効時ずーっとループする
            UnZipMonitor();

            Console.WriteLine(string.Format("監視終了"));
            return;
        }

        private static void CreateFolder( string path ) {
            if (!System.IO.Directory.Exists(path))
            {
                System.IO.Directory.CreateDirectory(path);
            }
        }

        /// <summary>
        /// モニタ有効時ずーっとループする
        /// </summary>
        static void UnZipMonitor()
        {
            Console.WriteLine(string.Format("{0}を監視中。終了するには何かキーを押してください。", Settings1.Default.自動解凍対象監視フォルダ));
            TimerCallback timerDelegate = new TimerCallback(_wTimer_Tick);
            _wTimer = new Timer(timerDelegate, null, 0, -1);
            Console.ReadKey();
        }


        /// <summary>
        /// 監視解凍
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        static void _wTimer_Tick(object o)
        {
            _wTimer.Change(-1, -1);
            UnZip(GetArgs(Directory.GetFiles(Settings1.Default.自動解凍対象監視フォルダ)));
            _wTimer.Change(Settings1.Default.監視間隔ミリ秒, -1);
        }

        /// <summary>
        /// 指定されたファイルたちを解凍します。
        /// </summary>
        /// <param name="targets"></param>
        private static void UnZip( IEnumerable<string> targets) 
        {
            foreach (string archive in targets)
            {
                UnZip(archive, _WORK_FOLDER);
                //解凍したファイルたちをリネームして元に戻す
                foreach (string unzipfile in Directory.GetFiles(_WORK_FOLDER))
                {
                    string restoreFilePath = Path.ChangeExtension(archive, Path.GetExtension(unzipfile));
                    if (!File.Exists(restoreFilePath))
                    {
                        File.Move(unzipfile, restoreFilePath);
                    }
                    else
                    {
                        Console.WriteLine(string.Format("既にファイルが存在するため無視しました。{0}", unzipfile));
                    }
                }
            }
        }
    

        /// <summary>
        /// 処理対象のパスを返します
        /// </summary>
        /// <param name="args"></param>
        /// <returns></returns>
        private static IEnumerable<string> GetArgs(string[] args)
        {
            return from x in args where x.LastIndexOf(".zip") > 0 select x;
        }

        /// <summary>
        /// 指定したパスから指定したパスへ解凍する
        /// </summary>
        private static bool UnZip(string sourcePath, string destPath ) 
        {
            //V版：ZIP解凍
            using (ZipFile zip = ZipFile.Read(sourcePath, Encoding.GetEncoding("Shift_JIS")))
            {
                //解凍時に既にファイルがあったら上書きする設定
                zip.ExtractExistingFile = ExtractExistingFileAction.OverwriteSilently;
                zip.FlattenFoldersOnExtract = true;
                //全て解凍する
                zip.ExtractAll(destPath);
            }
            //処理したら捨てる
            if (File.Exists(sourcePath))
            {
                File.Move(sourcePath, Path.Combine(_PROCESSED_FOLDER, Path.GetFileName(sourcePath)));
            }

            return true;
        }

    }
}
