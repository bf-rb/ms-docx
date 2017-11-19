module MsDocx
  # @author Alex Mercury
  class Paragraph

    # Nokogiri::XML::Element
    def initialize(xml_element)
      @xml_element = xml_element
    end

    def to_xml
      @xml_element.to_xml
    end

    def text
      @xml_element.xpath(@xml_element.path + '//w:r//w:t').map(&:content).join('')
    end


    def text=(value)

      text_node = @xml_element.xpath(@xml_element.path + '//w:r//w:t').first
      text_node.content = value if text_node

    end
  end
end
