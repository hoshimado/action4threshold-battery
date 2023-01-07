# action4threshold-battery

* タイトル
    * PowerShellでバッテリ残量を取得して、一定未満で且つユーザー応答が無ければOSをスリープ状態へ移行する

# 概要

本リポジトリは、次の目的を達成する方法を説明します。

* WindowsモバイルノートPCをバッテリー駆動で利用中に寝落ちした場合、自動的にスリープ状態に移行する

実現に使うツールは以下とします。いずれもWindowsに標準搭載されています。

* タスクスケジューラ
* PowerShell

実現だけすればよい、場合はスクリプトの解説を飛ばして「タスクスケジューラへの設定方法」の節を参照ください。

## 検証環境

「Windows 10」とします。


# 実現方法

次のようにして実現します。

1. タスクスケジューラで定期的に、監視用のPowerShellスクリプトを起動
2. PowerShellスクリプトでバッテリ残量を取得し、閾値以下なら「作業を継続するか？」のダイアログを表示
3. ダイアログは一定時間で自動クローズする設定とし、応答なしだった場合は「寝落ちしている」と判断してOSをスリープ状態へ移行

## 監視用スクリプトの作成

PowerShellでバッテリー残量を取得するには、WMIが提供するWin32_Batteryクラスを利用して、次のようにします。
ここでは、後で表示に使うので「残り使用可能時間」も合わせて取得しておきます。

```
Get-CimInstance -ClassName Win32_Battery | Select-Object -Property DeviceID, EstimatedChargeRemaining, EstimatedRunTime | Foreach-Object {
  # $DeviceID = $_.DeviceID

  # 残バッテリ量(100-0％) 
  $EstimatedChargeRemaining = $_.EstimatedChargeRemaining

  # 残り使用時間(分)
  $EstimatedRunTime = $_.EstimatedRunTime
}

```

一定の表示時間で自動的に閉じるダイアログをPowerShellで表示するには、次のようにします。
ここではWscriptのポップアップを利用する方法とし、「はい」「いいえ」を表示するようにvbYesNoe (=4) を指定します。

```
  $wsobj = new-object -comobject wscript.shell
  $msgText = "表示するメッセージ"
  $nSecondsToWait = 5 # 表示する秒数
  $titleText = "タイトル文字列"

  $result = $wsobj.popup($msgText, $nSecondsToWait, $titleText, "4")
```

OSをスリープ状態へ移行するには、.Net FrameworkのSetSuspendStateを利用して、次のようにします。
```
Add-Type -Assembly System.Windows.Forms;[System.Windows.Forms.Application]::SetSuspendState(‘Suspend’, $false, $false);
```

以上の内容を用いてスクリプトを作成すると、
本リポジトリの [message-sleep-battery.ps1](./message-sleep-battery.ps1) のようになります。


## タスクスケジューラへの設定方法

タスクスケジューラを起動して、次のようにタスクを設定します。

1. 「タスクの作成」を選択
2. タブ「全般」の「名前」に任意（例：message-battery-sleep等）の名称を入力
    * 他の項目はデフォルトのまま
3. タブ「トリガー」で

（以下、作成中）



# この方法を選んだ背景

「寝落ちした場合に、自動的にスリープ状態へ移行して欲しい」を実現しようと考えた場合に、
OS標準の「バッテリーの設定＞電源とスリープ＞次の時間が経過後、PCをスリープ状態にする」
の設定で対応可能な場合が多いです。

ただ、上記の設定では「ブラウザゲームを再生中」や「Youtubeを再生中」における
「寝落ち」への対処は出来ません（スリープ状態へ移行しません）。
上記の「時間が経過」は「アイドル状態」を意図していると推定され、
「ゲームや動画を再生中」は「Notアイドル状態」と扱われるため
対象外のようです。

なので、上記以外の方法でスリープ状態へ持って行く必要があります。

最初に検討したのは「The Marvellous Suspender」で、一定時間の操作が無い場合に
Chromeブラウザをスリープさせてしまえば「アイドル状態」になるのでは？というアプローチでした。
しかし、残念ながら「The Marvellous Suspender」は
「アクティブでは無いタブをスリープさせる」なので適用は出来きませんでした。
アクティブに成っているタブは、表示状態のままとなります。

https://cravelweb.com/gadget/pc/how-to-setting-the-marvellous-suspender-google-chrome-extension


次に、タスクスケジューラの「トリガー」にある「アイドル状態の時」を利用する方法を
検討しましたが、これも結局は「Notアイドル状態、と判定されている」ために
動作しませんでした。

https://4thsight.xyz/37398



バッテリー監視系のソフトで、そのような機能を有するものが無いかを検索したところ、
一応「バッテリー状態をトリガーとしてタスクを実行する」機能を有するソフトはありました。
しかし、あくまでサブの機能であり、今回の目的のためだけにソフト導入するのは、出来れば見送りたい、
と考えたため、こちらは保留としました。

https://www.gigafree.net/system/monitor/Battery-Mode.html


ならば、ということで、バッテリー状態を監視して条件を満たしたら、スリープさせる
アプリをいっそのこと作ってしまえば、、、とも考えましたが、ソコまで労力をかけるのも、
出来れば見送りたい、と考えて保留としました。


と言うわけで、スクリプトとタスクスケジューラを組み合わせて何とかできないか？
と考えた結果に、今回の実現方法に辿り着いた、というのが背景と成ります。


なお、タスクスケジューラから実行時に、ウィンドウを非表示にするにはVBS経由とする
必要があるので、スクリプト自体をvbsで書くことも考えましたが（とりあえずは作成済み）、
OSをスリープ状態にするにはPowerShellが最適である、と言う背景から、
スクリプト本体はPowerShellとする、という判断をしました。

https://neos21.net/blog/2022/02/23-01.html


PowerShell（だけでなくvbsも）は、今回に初めて書いたので、「ｘｘと書いた方がより適切」
などの指摘があれば、コメントいただけると助かります。





