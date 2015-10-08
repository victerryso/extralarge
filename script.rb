$time = Time.new

require 'nokogiri'
require 'zip'
require 'byebug'
require 'write_xlsx'

def time_stamp(string)
  time = Time.new - $time
  time = time.round(2)
  puts "#{time}s - #{string}"
end

time_stamp('Required Modules')

def open_zip_file(params)

  file_types  = [
    { name: :details,  path: /docProps\/core.*/         },
    { name: :strings,  path: /xl\/sharedStrings.*/      },
    { name: :sheets,   path: /xl\/worksheets\/sheet1.*/ },
    { name: :charts,   path: /xl\/charts\/chart.*/      },
    { name: :external, path: /xl\/external.*/           }
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
        cells = {}

        doc.css('c').each do |cell|
          formula = cell.css('f').text
          value   = cell.css('v').text
          next if formula == '' && value == ''

          cells[cell['r']] = {
            datatype:    cell['t'],
            formula:     formula,
            value:       value
          }
        end

        sheet = { name: entry.name, cells: cells }

        params[:sheets].push(sheet)

      when :charts
        chart = {
          title:  doc.css('c|chart > c|title').text,
          xTitle: doc.css('c|catAx').text,
          yTitle: doc.css('c|valAx').text
        }

        if chart[:xTitle] == ''
          chart[:xTitle] = doc.css('c|valAx')[0].text if doc.css('c|valAx')[0]
          chart[:yTitle] = doc.css('c|valAx')[1].text if doc.css('c|valAx')[1]
        end

        params[:charts].push(chart)

      when :external
        if doc.css('Relationship').first
          targets = doc.css('Relationship').map { |node| node.attr('Target') }
          params[:details][:external] ||= []
          params[:details][:external] += targets
        end

      end

    end

  end

end

# Initialize Template
def run_script(template_file, directory, output)

  template = {
    details: { excelFile: template_file },
    strings: [],
    sheets:  [],
    charts:  []
  }

  open_zip_file(template)

  template_coordinates = template[:sheets].first[:cells].keys

  # Initialize Student
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

    # Validate Cells
    student[:sheets].map do |sheet|
      # Reject Coordinates which were in template
      template_coordinates.each { |coordinate| sheet[:cells].delete(coordinate) }

      # Get Correct Values From DataType
      sheet[:cells].each do |coordinate, cell|

        case cell[:datatype]

        when 's'
          index = cell[:value].to_i
          string = student[:strings][index]
          cell[:value] = string

        when nil
          cell[:value] = cell[:value].to_f

        end

      end

    end

    student
  end

  time_stamp('Parsed Excel Files')

  # Check to see percentage of copied sheet
  students.each do |s1|

    matches = students.map do |s2|
      next if s1.eql?(s2)

      s1_cells = s1[:sheets].first[:cells].clone
      s2_cells = s2[:sheets].first[:cells].clone

      coordinate_matches = s1_cells.select { |coordinate, cell| s2_cells[coordinate] }
      cell_matches       = s1_cells.select { |coordinate, cell| cell == s2_cells[coordinate] }

      next if coordinate_matches.length == 0

      percent = cell_matches.length * 100 / coordinate_matches.length

      { percent: percent, copied: s2[:details][:excelFile] }
    end

    matches = matches.compact.sort_by { |m| m[:percent] }.last

    s1[:details][:percent] = matches && matches[:percent] || 0
    s1[:details][:copied]  = matches && matches[:copied]  || s1[:details][:excelFile]

    external = s1[:details][:external] || []
    s1[:details].delete(:external)
    s1[:details][:external] = external.join(' || ')
  end

  all_excel_files = []

  students.sort_by! do |student|
    -student[:details][:percent]
  end

  students.each do |student|
    unless all_excel_files.include?(student[:details][:copied])
      student[:details][:copied] = student[:details][:excelFile]
      all_excel_files << student[:details][:excelFile]
    end
  end

  students.sort_by! do |student|
    percent = student[:details][:percent] ? -(student[:details][:percent]) : 0
    [ percent, student[:details][:copied] ]
  end

  time_stamp('Collected Similarities')

  # Create Workbook
  workbook = WriteXLSX.new(output)
  worksheet = workbook.add_worksheet

  # Formating
  colors = [
    workbook.set_custom_color(21, 229, 115, 115),
    workbook.set_custom_color(22, 240, 98,  146),
    workbook.set_custom_color(23, 186, 104, 200),
    workbook.set_custom_color(24, 149, 117, 205),
    workbook.set_custom_color(25, 121, 134, 203),
    workbook.set_custom_color(26, 100, 181, 246),
    workbook.set_custom_color(27, 79,  195, 247),
    workbook.set_custom_color(28, 77,  208, 225),
    workbook.set_custom_color(29, 77,  182, 172),
    workbook.set_custom_color(30, 129, 199, 132),
    workbook.set_custom_color(31, 174, 213, 129),
    workbook.set_custom_color(32, 220, 231, 117),
    workbook.set_custom_color(33, 255, 241, 118),
    workbook.set_custom_color(34, 255, 213, 79 ),
    workbook.set_custom_color(35, 255, 183, 77 ),
    workbook.set_custom_color(36, 255, 138, 101),
    workbook.set_custom_color(37, 161, 136, 127),
    workbook.set_custom_color(38, 224, 224, 224),
    workbook.set_custom_color(39, 144, 164, 174)
  ]

  formats = colors.map { |color| workbook.add_format(:bg_color => color, :pattern => 0, :border => 0) }

  # Insert Headers and group chart/coordinate values
  details = students.map { |student| student[:details].keys }.flatten.uniq

  charts = []
  chart_keys = [:title, :xTitle, :yTitle]
  no_charts = students.map { |student| student[:charts].length }.max
  no_charts.times { charts.push(chart_keys) }

  all_charts = []
  all_charts = students.map { |student| (0..no_charts).map { |chart_index| student[:charts][chart_index] } }.uniq.flatten

  all_coordinates = students.map { |student| student[:sheets].first[:cells].keys }.flatten.uniq

  coordinates = {}
  all_coordinates.each do |coordinate|
    values = students.map { |student| student[:sheets].first[:cells][coordinate] }.uniq.compact
    coordinates[coordinate] = values
  end

  all_coordinates = all_coordinates[0...5000] if all_coordinates.length > 5000

  headers = {
    details:     details,
    charts:      charts.flatten,
    coordinates: all_coordinates.map { |coordinate| [ "#{coordinate} (Formula)", "#{coordinate} (Value)" ] }.flatten.uniq
  }

  headers.values.flatten.each_with_index { |header, col| worksheet.write(0, col, header.to_s) }

  # Insert Rows
  students.each_with_index do |student, row|

    # Insert Data (Details)
    row += 1
    student[:details].values.each_with_index { |detail, col| worksheet.write(row, col, detail) }

    # Insert Data (Charts)
    col = headers[:details].length

    student[:charts].each do |chart|
      match = all_charts.index(chart)
      format = match ? formats[match / no_charts] : nil

      chart.keys.each_with_index do |key|
        val = chart[key]
        match ? worksheet.write(row, col, val, format) : worksheet.write(row, col, val)
        col += 1
      end
    end

    # Insert Data (Coordinates)
    col = headers[:details].length + headers[:charts].length
    cells = student[:sheets].first[:cells]

    all_coordinates.each do |coordinate|
      cell = cells[coordinate]

      unless cell
        col += 2
        next
      end

      val = cell[:formula]
      worksheet.write(row, col, val)
      col += 1

      match = coordinates[coordinate].index(cell)
      format = match ? formats[match] : nil

      end

      val = cell[:value]
      match ? worksheet.write(row, col, val, format) : worksheet.write(row, col, val)
      col += 1
    end

  end

  workbook.close

  time_stamp('Workbook Created')

end



# 6.times do |n|
#   n += 1
#   run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
# end

n = 3
run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
