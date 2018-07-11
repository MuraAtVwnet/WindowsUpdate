概要
    Windows Update と再起動を自動実行する

スクリプト説明
    AutoWindowsUpdate.ps1
        C:\WindowsUpdateに置いて管理権限実行
        引数に「 Full 」を与えると全ての更新を適用する
        引数なしだと重要な更新のみを適用する
        適用完了するまで何度も自動再起動
        C:\WU_Log に実行ログ出力

    SubmitWU.ps1
        補助ツール
        AutoWindowsUpdate.ps1 を自動起動にセットして reboot する

        使いどころ
            いつまでも PowerShell プロンプト開いたままにしたくない
            リモートコンピューターに WU をかける

補足説明
    Auto 系のスクリプトは、実行ログに "=-=-=-=-=- Windows Update finished -=-=-=-=-=" が出力されるまで放置
    # そのまま使っていると、予告もなく再起動が入るので危険

    会話式の更新(Windows Defenderのパターンファイルとか)は適用されないので、最後に手動 Windows Update
    が必要(短時間で終わる)

    全ての更新を適用すると、.NET が最新バージョンまで上がるので要注意
    運用中の AP/TM は重要な更新のみ適用が原則

    UEFI セキュアブート環境(含むWindows Server 2012 R2 Gen 2 VM)だと、KB2962824 問題でロールバックが発
    生することがある。
    対策は、BitLocker モジュールをインストールする。(インストールのみ、BitLocker 構成不要)

        Add-WindowsFeature BitLocker -Restart

        http://support.microsoft.com/kb/2962824/ja より引用
        このエラーは、セキュリティ更新プログラム 2962824 のインストーラーが、BitLocker がインストール
        されていることを誤って予期するために発生します。

Windows Update が進まない時の対応

    30分以上待っても処理が進まない時は、Windows Update クライアントを更新すると改善します。
    (Windows Update クライアントは、Windows Update で配布されるのですが、初めての Windows Update だとこれが古いままなので手動で最新に更新します)

    Windows Update クライアントは以下キーワードで検索して最新版をダウンロードします。

        Windows Update クライアント site:support.microsoft.com

    ちなみに、これを書いているときの最新は以下の KB です。

        Windows 7 および Windows Server 2008 R2 用 Windows Update クライアント: 2016 年 3 月
        https://support.microsoft.com/ja-jp/kb/3138612
