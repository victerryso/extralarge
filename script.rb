$time = Time.new

require 'nokogiri'
require 'zip'
require 'byebug'
require 'write_xlsx'
require 'dentaku'

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

  Zip::File.open(params[:details][:filename]) do |file|

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
          formula = cell.css('f').text.downcase
          value   = cell.css('v').text.downcase
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
          title:  doc.css('c|chart > c|title').text.downcase,
          xTitle: doc.css('c|catAx').text.downcase,
          yTitle: doc.css('c|valAx').text.downcase,
          xRef:   doc.css('c|xVal c|f').text,
          yRef:   doc.css('c|yVal c|f').text
        }

        if chart[:xTitle] == ''
          chart[:xTitle] = doc.css('c|valAx')[0].text.downcase if doc.css('c|valAx')[0]
          chart[:yTitle] = doc.css('c|valAx')[1].text.downcase if doc.css('c|valAx')[1]
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
    details: { filename: template_file },
    strings: [],
    sheets:  [],
    charts:  []
  }

  open_zip_file(template)

  template_coordinates = template[:sheets].first[:cells].keys

  # Initialize Student
  students = Dir.entries(directory)

  students.select! { |student| /^[^~].*xlsx$/ === student}

  students.map! do |student|

    student_file = directory + student

    student = {
      details: { filename: student_file },
      strings: [],
      sheets:  [],
      charts:  []
    }

    open_zip_file(student)

    # Validate Cells
    student[:sheets].map do |sheet|

      # Reject Coordinates which were in template
      # template_coordinates.each { |coordinate| sheet[:cells].delete(coordinate) }

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

  # Marking
  if template_file == 'templates/template3.xlsx'

    # Evaluate Answers - Template 3
    questions = [
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D18' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D19' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D20' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D21' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D22' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D23' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D24' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D25' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D26' },
      { number: '2A', type: :fill,      marks: 0.1,  primary: 'D27' },

      { number: '2B', type: :condition, marks: 0.25, primary: 'L18 * 21 = M18',    secondary: 'M18'   },
      { number: '2B', type: :condition, marks: 0.25, primary: 'L19 * 21 = M19',    secondary: 'M19'   },
      { number: '2B', type: :condition, marks: 0.25, primary: 'L20 * 21 = M20',    secondary: 'M20'   },
      { number: '2B', type: :condition, marks: 0.25, primary: 'L21 * 21 = M21',    secondary: 'M21'   },

      { number: '2C', type: :regex,     marks: 1,    primary: /\$M\$18\:\$M\$21/i, secondary: :yRef   },

      { number: '2D', type: :regex,     marks: 0.34, primary: /VO2/i,              secondary: :title  },
      { number: '2D', type: :regex,     marks: 0.33, primary: /(heart rate|hr)/i,  secondary: :yTitle },
      { number: '2D', type: :regex,     marks: 0.33, primary: /bpm/i,              secondary: :yTitle },

      { number: '2E', type: :regex,     marks: 1,    primary: /\$F\$18\:\$G\$21/i, secondary: :xRef   },

      { number: '7',  type: :equal,     marks: 1,    primary: 'B',                 secondary: 'B92'   }
    ]

    students.each do |student|
      calculator = Dentaku::Calculator.new

      student[:sheets].first[:cells].each do |coordinate, cell|
        store = {}
        store[coordinate] = cell[:value]
        calculator.store(store)
      end

      student[:marks] = questions.map do |question|
        correct = case question[:type]

        when :condition
          calculator.evaluate(question[:primary])

        when :equal
          cell = student[:sheets].first[:cells][question[:secondary]]
          cell && question[:primary] == cell[:value]

        when :regex
          titles = student[:charts].map do |chart|
            question[:primary] === chart[question[:secondary]]
          end

          titles.select! { |t| t }
          titles.compact.any?

        when :fill
          student[:sheets].first[:cells][question[:primary]]

        end
        correct ? question[:marks] : 0
      end

      student[:marks] << student[:marks].inject(:+)

    end
  end

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

      { percent: percent, copied: s2[:details][:filename] }
    end

    matches = matches.compact.sort_by { |m| m[:percent] }.last

    s1[:details][:percent] = matches ? matches[:percent] : 0
    s1[:details][:copied]  = matches ? matches[:copied]  : s1[:details][:filename]

    external = s1[:details][:external] || []
    s1[:details].delete(:external)
    s1[:details][:external] = external.join(' || ')
  end

  all_excel_files = []

  students.sort_by! { |student| -student[:details][:percent] }

  students.each do |student|
    unless all_excel_files.include?(student[:details][:copied])
      student[:details][:copied] = student[:details][:filename]
      all_excel_files << student[:details][:filename]
    end
  end

  students.sort_by! { |student| [ -student[:details][:percent], student[:details][:copied] ] }

  students.sort_by! do |student|
    sum1 = students.select { |s| s[:details][:copied] == student[:details][:copied] }.length
    sum2 = students.select { |s| s[:details][:copied] == student[:details][:copied] }.map { |s| s[:details][:percent] }.inject(:+)
    [ -sum1, -sum2, student[:details][:copied], -student[:details][:percent] ]
  end

  time_stamp('Collected Similarities')

  # Flagging Students
  students.each do |student|
    flags = [
      # { type: :external },
      { type: :cell, coordinate: 'B83'}

    ]

    flags.map do |flag|
      color = case flag[:type]

      when :external
        student[:external]

      when :filename
        students.find do |s|
          s[:details][:filename] == student[:details][:filename] && !s.eql?(student)
        end

      when :cell
        cell = student[:sheets].first[:cells][flag[:coordinate]]
        students.find do |s|
          s_cell = s[:sheets].first[:cells][flag[:coordinate]]

          !cell.nil? && cell.eql?(s_cell) && !s.eql?(student)
        end

      when :chart
        chart = property
        students.find do |s|
          s[:charts].first[:title] == student[:charts].first[:title] && !s.eql?(student)
        end

      when :easter
        false

      end
      p [ color[:sheets].first[:cells][flag[:coordinate]], student[:sheets].first[:cells][flag[:coordinate]] ] if color
      # p color
      p !!color
    end
  end

end



# 6.times do |n|
#   n += 1
#   run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
# end

n = 3
run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
