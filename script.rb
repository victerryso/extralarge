$time = Time.new

require 'nokogiri'
require 'zip'
require 'byebug'
require 'write_xlsx'

def time_stamp(string)
  puts "#{Time.new - $time} - #{string}"
end

time_stamp('Done Requiring')

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
          yTitle: doc.css('c|valAx').text,
          xRef:   doc.css('c|xVal c|f').text,
          yRef:   doc.css('c|yVal c|f').text
        }

        if chart[:xTitle] == ''
          chart[:xTitle] = doc.css('c|valAx')[0].text if doc.css('c|valAx')[0]
          chart[:yTitle] = doc.css('c|valAx')[1].text if doc.css('c|valAx')[1]
        end

        if chart[:xRef] == '' then chart[:xRef] = doc.css('c|cat c|f').text end
        if chart[:yRef] == '' then chart[:yRef] = doc.css('c|val c|f').text end

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
      p sheet[:cells].keys.length
      # Reject Coordinates which were in template
      template_coordinates.each { |coordinate| sheet[:cells].delete(coordinate) }
      p sheet[:cells].keys.length

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

        cell.delete(:datatype)
      end

    end

    student
  end

  time_stamp('Done Grabbing Excel Files')

  # Check to see percentage of copied sheet
  students.each do |s1|

    matches = students.map do |s2|
      next if s1.eql?(s2)

      s1_cells = s1[:sheets].first[:cells].clone
      s2_cells = s2[:sheets].first[:cells].clone

      coordinate_matches = s1_cells.select { |coordinate, cell| s2_cells[coordinate] }
      cell_matches      = s1_cells.select { |coordinate, cell| cell == s2_cells[coordinate] }

      next if coordinate_matches.length == 0

      percent = cell_matches.length * 100 / coordinate_matches.length

      { percent: percent, copied: s2[:details][:excelFile] }
    end

    matches = matches.compact.sort_by { |m| m[:percent] }.last

    s1[:details][:percent] = matches[:percent] || 0
    s1[:details][:copied] = matches[:copied] || s1[:details][:excelFile]

    external = s1[:details][:external] || []
    s1[:details].delete(:external)
    s1[:details][:external] = external.join(' || ')
  end

  time_stamp('Done Percentage of Likeness')

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

  time_stamp('Done Checking Similar Files')

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

  # Insert Headers (Details)
  details = students.first[:details].keys
  details.each_with_index { |detail, index| worksheet.write(0, index, detail) }

  # Insert Headers (Charts)
  headers = students.first[:charts].map { |chart| chart.keys }.flatten
  headers.each_with_index { |header, index| worksheet.write(0, details.length + index, header) }

  byebug

  # Insert Headers (Coordinates)
  all_coordinates = students.map { |student| student[:sheets].first[:cells].keys }.flatten.uniq
  all_coordinates -= template_coordinates
  all_coordinates.sort_by! { |coordinate| [ coordinate.gsub(/\D/, '').to_i, coordinate.gsub(/\d/, '') ] }

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

    # Insert Data (Details)
    row = index + 1
    student[:details].values.each_with_index { |detail, i| worksheet.write(row, i, detail) }

    # Insert Data (Charts)
    # student[:charts].each_with_index do |chart, chart_index|
    #   col = details.length + chart_index * 5
    #   chart.values.each_with_index  |value, value_index| worksheet.write(row, col + value_index, value) }
    # end

    # student[:charts].each_with_index do |chart, chart_index|
    #   chart.keys.each_with_index do |key, key_index|
    #     matches = students.map { |s| s[:charts] }.flatten
    #     matches = matches.select { |m| m[key].eql?(chart[key]) }.length
    #     matches = matches > formats.length ? formats.length - 1 : matches
    #
    #     col = details.length + chart_index * 5 + key_index
    #     val = chart[key]
    #     # chart.values.each_with_index { |value, value_index| worksheet.write(row, col + value_index, value) }
    #
    #     if matches
    #       worksheet.write(row, col, val, formats[matches])
    #     else
    #       worksheet.write(row, col, val)
    #     end
    #
    #   end
    # end

    student[:charts].each_with_index do |chart, chart_index|
      chart.keys.each_with_index do |key, key_index|
        # matches = students.map { |s| s[:charts] }.flatten.map { |m| m[key] }.uniq
        # matches = matches.index(chart[key])

        col = details.length + chart_index * 5 + key_index
        val = chart[key]

        # if matches
        #   worksheet.write(row, col, val, formats[matches])
        # else
          worksheet.write(row, col, val)
        # end

      end
    end


    # Insert Data (Coordinates)
    cells = student[:sheets].first[:cells]
    cells.each do |cell|
      # Insert Data (Formulas)
      col = 2 * all_coordinates.find_index(cell[:coordinates]) + details.length + headers.length
      val = cell[:formula]

      worksheet.write(row, col, val)

      # Insert Data (Values)
      col += 1
      val = cell[:value]

      # Format Cells
      # matches = students.map { |s| s[:sheets].first[:cells].find { |c| c.eql?(cell) && !s.eql?(student) } }.compact.length
      #
      # if matches >= formats.length
      #   worksheet.write(row, col, val, formats[formats.length - 1])
      # elsif matches > 0
      #   worksheet.write(row, col, val, formats[matches - 1])
      # else
      #   worksheet.write(row, col, val)
      # end

      # Format Cells v2
      # all_values = students.map { |s| s[:sheets].first[:cells].find { |c| c[:coordinates] == cell[:coordinates] } }
      # all_values = all_values.compact.select { |v| all_values.select { |vv| v.eql?(vv) }.length > 1 }
      # value_index = all_values.uniq.any? ? all_values.uniq.find_index(cell) : nil
      #
      # if value_index
      #   worksheet.write(row, col, val, formats[value_index])
      # else
        worksheet.write(row, col, val)
      # end


    end
  end

  workbook.close

  time_stamp('Done Creating Workbook')

end



# 6.times do |n|
#   n += 1
#   run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
# end

n = 3
run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
