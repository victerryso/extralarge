require 'nokogiri'
require 'zip'
require 'byebug'
require 'write_xlsx'

def open_zip_file(params)

  file_types  = [
    { name: :details, path: /docProps\/core.*/        },
    { name: :strings, path: /xl\/sharedStrings.*/     },
    { name: :sheets,  path: /xl\/worksheets\/sheet.*/ },
    { name: :charts,  path: /xl\/charts\/chart.*/     }
  ]


  Zip::File.open(params[:details][:excelFile]) do |file|

    # Handle entries one by one
    file.each do |entry|

      # Grab File Type
      file_type = file_types.find { |type| type[:path] === entry.name }

      # Skip other file types
      next unless file_type

      # Read into memory
      content = entry.get_input_stream.read
      doc = Nokogiri::XML(content)

      # Depends on the file type
      case file_type[:name]

      when :details
        params[:details][:name] = doc.css('cp|lastModifiedBy').text
        params[:details][:date] = doc.css('dcterms|modified').text

      when :strings
        params[:strings] = doc.css('si').map { |si| si.text }

      when :sheets
        cells = doc.css('c').map do |cell|
          {
            coordinates: cell['r'],
            datatype:    cell['t'],
            formula:     cell.css('f').text,
            value:       cell.css('v').text
          }
        end

        cells.reject! { |cell| cell[:formula] == '' && cell[:value] == '' }
        sheet = { name: entry.name, cells: cells }

        params[:sheets].push(sheet)

      when :charts
        chart = {
          title:  doc.css('c|chart > c|title').text,
          xTitle: doc.css('c|catAx').text,
          yTitle: doc.css('c|valAx').text,
          xRef:   doc.css('c|xVal c|f').text,
          yRef:   doc.css('c|yVal c|f').text
        }

        if chart[:xTitle] == ''
          chart[:xTitle] = doc.css('c|valAx')[0].text
          chart[:yTitle] = doc.css('c|valAx')[1].text
        end

        if chart[:xRef] == '' then chart[:xRef] = doc.css('c|cat c|f').text end
        if chart[:yRef] == '' then chart[:yRef] = doc.css('c|val c|f').text end

        params[:charts].push(chart)

      end

    end

    # # Find specific entry
    # entry = file.glob('*.csv').first
    # puts entry.get_input_stream.read
  end

end

# Initialize Template
template = {
  details: { excelFile: 'template.xlsx' },
  strings: [],
  sheets:  [],
  charts:  []
}

open_zip_file(template)

template[:sheets].reject! { |sheet| sheet[:cells].length < 10 }

# Initialize Student
directory = 'students/'
students = Dir.entries(directory)
students.select! { |student| /^[^~].*xlsx$/ === student}
students.map! { |student| directory + student }

students.map! do |student_file|
  student = {
    details: { excelFile: student_file },
    strings: [],
    sheets:  [],
    charts:  []
  }

  open_zip_file(student)

  student[:sheets].reject! { |sheet| sheet[:cells].length < 10 }

  # Validate Cells
  student[:sheets].each_with_index.map do |sheet, index|

    # Reject Coordinates which were already there
    sheet[:cells].reject! do |cell|
      template[:sheets][index][:cells].find do |tCell|
        tCell[:coordinates] == cell[:coordinates]
      end
    end

    # Get Correct Values From DataType
    sheet[:cells].map! do |cell|
      case cell[:datatype]
      when 's'
        index = cell[:value].to_i
        string = student[:strings][index]
        cell[:value] = string

      when nil
        cell[:value] = cell[:value].to_f

      end

      cell.delete(:datatype)
      cell
    end

    sheet[:cells].sort_by { |cell| cell[:coordinates] }

    # sheet[:cells].map { |cell| p cell }
  end

  student
end

# Create Workbook
workbook = WriteXLSX.new('output.xlsx')
worksheet = workbook.add_worksheet

# Formating
blues = [
  workbook.set_custom_color(11, 227, 242, 253),
  workbook.set_custom_color(12, 187, 222, 251),
  workbook.set_custom_color(13, 144, 202, 249),
  workbook.set_custom_color(14, 100, 181, 246),
  workbook.set_custom_color(15, 66,  165, 245),
  workbook.set_custom_color(16, 33,  150, 243),
  workbook.set_custom_color(17, 30,  136, 229),
  workbook.set_custom_color(18, 25,  118, 210),
  workbook.set_custom_color(19, 21,  101, 192),
  workbook.set_custom_color(20, 13,  71,  161)

]

formats = blues.map { |blue| workbook.add_format(:bg_color => blue, :pattern => 0, :border => 0) }

# Insert Headers (Details)
details = students.first[:details].keys
details.each_with_index { |detail, index| worksheet.write(0, index, detail) }

# Insert Headers (Charts)
headers = students.first[:charts].each_with_index.map { |chart, index| chart.keys.map { |key| "#{key} (#{index})"} }.flatten
headers.each_with_index { |header, index| worksheet.write(0, details.length + index, header) }

# # Insert Headers (Coordinates)
all_coordinates = students.map { |student| student[:sheets].first[:cells].map { |cell| cell[:coordinates] } }.flatten.uniq
all_coordinates.each_with_index do |coordinate, index|
  row = 0
  col = 2 * index + details.length + headers.length
  val = "#{coordinate} (Formula)"

  worksheet.write(row, col, val)

  col += 1
  val = "#{coordinate} (Value)"

  worksheet.write(row, col, val)
end

# Insert Rows
students.each_with_index do |student, index|
  row = index + 1
  student[:details].values.each_with_index { |detail, i| worksheet.write(row, i, detail) }

  matches = students.find { |s| !(s[:sheets].first[:cells] - student[:sheets].first[:cells]).any? && !s.eql?(student)}

  if matches
    worksheet.write(row, 2, student[:excelFile], formats[-1])
  else
    worksheet.write(row, 2, student[:excelFile])
  end

  student[:charts].each_with_index do |chart, chart_index|
    col = details.length + chart_index * 5
    chart.values.each_with_index { |value, value_index| worksheet.write(row, col + value_index, value) }
  end

  cells = student[:sheets].first[:cells]
  cells.each do |cell|
    col = 2 * all_coordinates.find_index(cell[:coordinates]) + details.length + headers.length
    val = cell[:formula]

    worksheet.write(row, col, val)

    col += 1
    val = cell[:value]

    matches = students.map { |s| s[:sheets].first[:cells].find { |c| c.eql?(cell) && !s.eql?(student) } }.compact.length

    if matches >= formats.length
      worksheet.write(row, col, val, formats[formats.length - 1])
    elsif matches > 0
      worksheet.write(row, col, val, formats[matches - 1])
    else
      worksheet.write(row, col, val)
    end


  end
end

workbook.close
