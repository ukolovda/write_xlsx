# coding: utf-8
# frozen_string_literal: true

#
# XMLWriterSimple
#
require 'stringio'

module Writexlsx
  module Package
    class XMLWriterSimple
      XMLNS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

      def initialize
        @io = []
        # @io = StringIO.new
        # Will allocate new string once, then use allocated string
        # Key is tag name
        # Only tags without attributes will be cached
        @tag_start_cache = {}
        @tag_end_cache = {}
      end

      def set_xml_writer(filename = nil)
        @filename = filename
      end

      def xml_decl(encoding = 'UTF-8', standalone = true)
        str = %(<?xml version="1.0" encoding="#{encoding}" standalone="#{standalone ? "yes" : "no"}"?>\n)
        io_write([str])
        str
      end

      def tag_elements(tag, attributes = nil)
        start_tag(tag, attributes)
        yield
        end_tag(tag)
      end

      def tag_elements_str(tag, attributes = nil)
        start_tag_str(tag, attributes) +
          yield +
          end_tag_str(tag)
      end

      def start_tag(tag, attr = nil)
        io_write start_tag_arr(tag, attr)
      end

      def start_tag_str(tag, attr = nil)
        start_tag_arr(tag, attr).join
      end

      def start_tag_arr(tag, attr)
        if attr.nil? || attr.empty?
          @tag_start_cache[tag] ||= ["<", tag, ">"]
        else
          # puts tag, attr
          attr.map! { |a| key_val_attr(a) }
          attr.prepend tag
          attr.prepend "<"
          attr << ">"
          attr
          # ["<", tag, key_vals(attr), ">"]
        end
      end

      def end_tag(tag)
        io_write end_tag_arr(tag)
      end

      def end_tag_str(tag)
        end_tag_arr(tag).join
      end

      def end_tag_arr(tag)
        @tag_end_cache[tag] ||= ["</", tag, ">"]
      end

      def empty_tag(tag, attr = nil)
        str = "<#{tag}#{key_vals(attr)}/>"
        io_write(str)
      end

      def empty_tag_encoded(tag, attr = nil)
        io_write(empty_tag_encoded_str(tag, attr))
      end

      def empty_tag_encoded_str(tag, attr = nil)
        "<#{tag}#{key_vals(attr)}/>"
      end

      def data_element(tag, data, attr = nil)
        tag_elements(tag, attr) { io_write(escape_data(data)) }
      end

      #
      # Optimised tag writer ?  for shared strings <si> elements.
      #
      def si_element(data, attr)
        tag_elements('si') { data_element('t', data, attr) }
      end

      #
      # Optimised tag writer for shared strings <si> rich string elements.
      #
      def si_rich_element(data)
        io_write("<si>#{data}</si>")
      end

      def characters(data)
        io_write(escape_data(data))
      end

      def crlf
        io_write("\n")
      end

      def close
        File.open(@filename, "wb:utf-8:utf-8") { |f| f << string } if @filename
        # @io.close
      end

      def string
        # @io.string
        @io = @io.join.force_encoding("UTF-8")
      end

      def io_write(str)
        if str.is_a?(Array)
          # @io += str
          str.each do |s|
            @io << s
          end
        else
          @io << str
        end
        str
      end

      private

      def key_val(key, val)
        %( #{key}="#{val}")
      end

      # attr is an Array or string
      def key_val_attr(attr)
        attr.is_a?(Array) ?
          key_val(attr.first, escape_attributes(attr.last)) :
          attr
      end

      def key_vals(attribute)
        if attribute
          case attribute.length
            when 0
              ''
            when 1
              key_val_attr(attribute[0])
            when 2
              key_val_attr(attribute[0])+key_val_attr(attribute[1])
            else
              attribute.inject('') do |str, attr|
                # attr can be array with 2 items, or string (constant, i.e. ' tag="value"')
                str + key_val_attr(attr)
              end
          end
        end
      end

      def escape_attributes(str = '')
        return str unless str.respond_to?(:match) && str =~ /["&<>\n]/

        str
          .gsub(/&/, "&amp;")
          .gsub(/"/, "&quot;")
          .gsub(/</, "&lt;")
          .gsub(/>/, "&gt;")
          .gsub(/\n/, "&#xA;")
      end

      def escape_data(str = '')
        if str.respond_to?(:match) && str =~ /[&<>]/
          str.gsub(/&/, '&amp;')
             .gsub(/</, '&lt;')
             .gsub(/>/, '&gt;')
        else
          str
        end
      end
    end
  end
end
