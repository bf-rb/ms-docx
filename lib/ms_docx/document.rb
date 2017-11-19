module MsDocx
  # @author Alex Mercury
  # @example
  #   doc = MsDocx::Document.open('1.docx')
  #
  #   puts doc.paragraphs.first.text
  #   # => 'some text'
  #
  class Document

    CONTENT_ENTRY = 'word/document.xml'

    class << self

      # Open .docx file
      # @param file_path[String] file path
      # @return [MsDocx::Document] document object
      def open(file_path)
        MsDocx::Document.new(file_path)
      end

    end

    def initialize(file_path, *args)
      if File.exist? file_path
        @file_path = file_path
        @temp_file_path = file_path.sub(/\.docx/, '~temp.docx')
        @docx_zip = Zip::File.open(file_path)
      else
        raise MsDocx::Errors::FileNotExist, "File '#{file_path}' doesn't exist ❗"
      end

      @doc = @docx_zip.find_entry(CONTENT_ENTRY).get_input_stream.read

      @xml = Nokogiri::XML::Document.parse @doc

      @paragraphs = []

      @xml.xpath('//w:body//w:p').each do |w_p|
        @paragraphs << MsDocx::Paragraph.new(w_p)
      end


      # @xml.xpath('//w:p//w:r//w:t').each do |p|
      #
      #
      #   puts p.content
      #
      #
      #   p.content = 'test'
      #
      #
      # end
      #

      #
      #
      # File.open(new_file, 'wb') {|f| f.write(buffer.string) }
    end

    def paragraphs
      @paragraphs
    end

    def each_paragraph(&block)
      @paragraphs.each do |paragraph|
        block.call(paragraph)
      end






      # block.call(self)


    end

    def to_text
      @paragraphs.map(&:text).join("\n")
    end

    def to_xml

      body = @xml.at('//w:body')

      body.children = ''

      @paragraphs.each do |paragraph|
        body.add_child(paragraph.to_xml)
      end

      @xml.at('//w:document').children = body

      @xml.to_xml
    end

    def save_file(file_path)

      @buffer = Zip::OutputStream.write_buffer do |out|
        @docx_zip.entries.each do |e|
          if e.name == CONTENT_ENTRY

            # doc = get_xml_doc_from_stream(e.get_input_stream.read)

            # doc.xpath('//w:body//w:p').each do |pt|

              # pt.xpath(pt.path + '//w:t').

              # p pt.to_xml

              # block.call(pt)

              # p pt.to_xml

              # w_r_list = pt.xpath(pt.path + '//w:r')

              # if w_r_list.length > 1
              #   w_r_list[1..-1].each{|pp| pp.remove }

                # p pt.to_xml
                # p w_r_list.to_xml
              # end

              # pt

              # p w_r_list.to_xml


              # w_r_list.first.xpath(w_r_list.first.path + '//w:t').map{|pp| block.call(pp)} if w_r_list.first


              # p nil
              # p nil
              # p nil


            # end


            # p doc.to_xml

            out.put_next_entry(CONTENT_ENTRY)
            out.write self.to_xml
          else
            out.put_next_entry(e.name)
            e.write_to_zip_output_stream out
          end
        end
      end


      File.open(file_path, 'wb') {|f| f.write(@buffer.string) }
    end

    private

      def get_xml_doc_from_stream(stream)
        Nokogiri::XML::Document.parse stream
      end

  end
end
