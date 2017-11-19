doc = MsDocx::Document.open('1.docx')

puts doc.paragraphs.first.text
