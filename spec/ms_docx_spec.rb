RSpec.describe MsDocx do

  xcontext do

    it 'test' do

    end

  end

  it 'Test' do
    file_path = __dir__ + '/docx_files/'

    doc = MsDocx::Document.open(file_path + 'test1.docx')

    # puts doc.paragraphs[0].to_xml

    # doc.each_paragraph do |par|
    #
    #   par.text = 'Test'
    #
    # end

    doc.paragraphs[0].text = 'Hello'

    puts doc.to_xml

    # doc.file_data

    # doc.save_file(file_path + 't2.docx')
  end

end
