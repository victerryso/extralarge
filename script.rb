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

  # Flagging Students
  students.each do |student|

    # Test Flags
    flags = [
      { type: :external, color: :red    },
      { type: :filename, color: :red    },
      { type: :author,   color: :orange },
      { type: :cell,     color: :red,    coordinate: 'B83' },
      { type: :chart,    color: :orange, property: :title  }
    ]

    # Show Colors
    matches = flags.map do |flag|
      match = case flag[:type]

      when :external
        student[:details][:external]

      when :filename
        filename = student[:details][:filename]
        students.reject { |s| s.eql?(student) }.count { |s| filename == s[:details][:filename] } > 0

      when :author
        name = student[:details][:name]
        students.reject { |s| s.eql?(student) }.count { |s| name == s[:details][:name] } > 0

      when :cell
        cell = student[:sheets].first[:cells][flag[:coordinate]]
        students.reject { |s| s.eql?(student) }.count { |s| cell.eql?(s[:sheets].first[:cells][flag[:coordinate]]) } > 0

      when :chart
        chart = student[:charts].first
        property = chart ? chart[flag[:property]] : nil
        students.reject { |s| s.eql?(student) }.map { |s| s[:charts] }.flatten.count { |c| c[flag[:property]] == property } > 0

      # TODO - Insert Text Somewhere and Find it Here
      # when :easter

      end

      match ? flag[:color] : :green
    end

    p matches
  end

  # Marking
  return
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
        titles = student[:charts].map { |chart| question[:primary] === chart[question[:secondary]] }

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





# 6.times do |n|
#   n += 1
#   run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
# end

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



n = 3
run_script("templates/template#{n}.xlsx", "worksheets/worksheets#{n}/", "outputs/output#{n}.xlsx")
