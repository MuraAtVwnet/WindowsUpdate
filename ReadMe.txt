概要
    Windows Update と再起動を Windows Update が完了するまで自動実行します

    Windows Server 2008 R2 / Windows 7 以降、PowerShell 3.0 以降で使えます
        (Windows Server 2016 / Windows 10 で動作確認)

スクリプト説明
    AutoWindowsUpdate.ps1
        自動 Windows Update スクリプト本体
        更新プログラム適用完了するまで何度も自動再起動

        C:\WindowsUpdateに置いて管理権限実行
        C:\WU_Log に実行ログ出力

        引数に「 Full 」を与えると全ての更新を適用します
        引数なし or 「 Minimum 」を与えると重要な更新のみを適用します

        Build Update 時に boot loop に陥ることがあるので、Build Update が含まれる場合は -ConsiderationBU オプションを指定してください
        (稼働時間が短い時はこのスクリプトを実行しないようにします)

    SetupSchedule-DaylyWU.ps1
        AutoWindowsUpdate.ps1 を毎日指定時刻に実行するスケジュールを登録します

        引数
            -StartTime : 開始時刻 HH:MM の 24h 表記
            -WuOption  : AutoWindowsUpdate.ps1 のオプションと同じ
            -ConsiderationBU : Build Update 時に boot loop に陥らないようにする

        例
            SetupSchedule-DaylyWU.ps1 -StartTime 04:00 -WuOption Full -ConsiderationBU
            毎日 4:00 AM に Full オプションで AutoWindowsUpdate.ps1 を実行するスケジュールを登録します

    SetupSchedule-MonthlyWU.ps1
        AutoWindowsUpdate.ps1 を毎月一度 Windows Update 日 + ディレイ日に実行するスケジュールを登録します

        引数
            -StartTime : 開始時刻 HH:MM の 24h 表記
            -WuDelay   : Windows Update 日(第2火曜日の翌日)からのディレイ日
            -WuOption  : AutoWindowsUpdate.ps1 のオプションと同じ
            -ConsiderationBU : Build Update 時に boot loop に陥らないようにする

        例
            SetupSchedule-MonthlyWU.ps1 -StartTime 04:00 -WuDelay 3 -WuOption Full -ConsiderationBU
            毎月の Windows Update 日から3日経過した 4:00 AM に Full オプションで AutoWindowsUpdate.ps1 を実行するスケジュールを登録します

    RemoveLog.ps1
        内部処理用(単独でこのスクリプトを使うことは想定していません)
        SetupSchedule-DaylyWU.ps1 / SetupSchedule-MonthlyWU.ps1 がスケジュールに登録するスクリプト
        過去ログ(3か月前)を削除します

    GoWU.ps1
        内部処理用(単独でこのスクリプトを使うことは想定していません)
        SetupSchedule-MonthlyWU.ps1 がスケジュールに登録するスクリプト
        毎月の Windows Update 日から引数で指定されたディレイ日であれば AutoWindowsUpdate.ps1 を実行します

    SubmitWU.ps1
        補助ツール
        AutoWindowsUpdate.ps1 を自動起動にセットして reboot します

        使いどころ
            いつまでも PowerShell プロンプト開いたままにしたくない場合
            リモートコンピューターに WU をかける場合

補足説明
    Auto 系のスクリプトは、実行ログに "=-=-=-=-=- Windows Update finished -=-=-=-=-=" が出力されるまで放置
    # そのまま使っていると、予告もなく再起動が入るので危険です

    会話式の更新は適用されないので、必要に応じて手動 Windows Update して下さい(短時間で終わるハズ)

    Full オプションで更新を適用すると、.NET が最新バージョンまで上がるので .NET バージョンに依存している環境は要注意です

    UEFI セキュアブート環境(含むWindows Server 2012 R2 Gen 2 VM)だと、KB2962824 問題でロールバックが発生することがあるので、BitLocker モジュールをインストールします。
    (インストールのみ、BitLocker 構成はしません)

        http://support.microsoft.com/kb/2962824/ja より引用
        このエラーは、セキュリティ更新プログラム 2962824 のインストーラーが、BitLocker がインストールされていることを誤って予期するために発生します。

Windows Update が進まない時の対応

    30分以上待っても処理が進まない時は、Windows Update クライアントを更新すると改善します。
    (Windows Update クライアントは、Windows Update で配布されるのですが、初めての Windows Update だとこれが古いままなので手動で最新に更新します)

    Windows Update クライアントは以下キーワードで検索して最新版をダウンロードします。

        Windows Update クライアント site:support.microsoft.com

    ちなみに、これを書いているときの最新は以下の KB です。

        Windows 7 および Windows Server 2008 R2 用 Windows Update クライアント: 2016 年 3 月
        https://support.microsoft.com/ja-jp/kb/3138612
