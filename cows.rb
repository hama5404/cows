# Gemfileのgemを一括require
require 'bundler'
Bundler.require

# EXCELファイルをOPEN
excel_path = './infiles/members.xlsx'
book = Roo::Spreadsheet.open(excel_path)

# シート名を指定する(必須ではない)
sheet = book.sheet('2021年名簿')

# 存在する最終列、最終行の確認
puts "最後の列番号: #{sheet.last_column}  最後の行番号: #{sheet.last_row}"
# 'B3'セルを参照する場合はsheet.cell(3, 2) or sheet.cell(3, 'B')

# シートの行数分でループ。先頭行はヘッダーとしてスキップ
(2..sheet.last_row).each do |idx|

  # PDF初期化
  # A6(はがきサイズ) 横向き　縁なし
  pdf = Prawn::Document.new(
    page_size: "A6", page_layout: :landscape,
    top_margin: 0, bottom_margin: 0, left_margin: 0, right_margin: 0)

  # 日本語フォントの読み込み
  pdf.font "./infiles/ipam.ttf"

  # 背景画像 縁なしで配置
  pdf.image("./infiles/background.png", at: [0,pdf.cursor])

  # 描画エリアを定義
  # cursorはまだ上
  pdf.bounding_box([20, pdf.cursor-10], width: 500) do
    # 'B3'セルを参照する場合はsheet.cell(3, 2) or sheet.cell(3, 'B')

    # 名前
    pdf.font_size(20)
    sei = sheet.cell(idx, "A")
    mei = sheet.cell(idx, "B")
    puts  " -- [#{idx}] #{sei} #{mei} さん"
    pdf.text "#{sei} #{mei} さん", color: "ffffff"

    # アケオメ
    pdf.move_down 15
    pdf.font_size(24)
    greet = "あけましておめでとうございます"
    pdf.text greet, color: "ffffff"
  end

  # うし
  pdf.bounding_box([300, pdf.cursor-80], width: 500) do
    pdf.font_size(80)
    pdf.text "丑",  color: "ff75b1"
  end

  # PDF保存  ファイル名 "outfiles/last_first.pdf"
  file_name = "./outfiles/#{sheet.cell(idx,"C")}_#{sheet.cell(idx,"D")}.pdf"
  pdf.render_file(file_name)

end