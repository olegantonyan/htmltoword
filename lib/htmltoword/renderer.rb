module Htmltoword
  class Renderer
    class << self
      def send_file(context, filename, options = {})
        new(context, filename, options).send_file
      end
    end

    def initialize(context, filename, options)
      @word_template = options[:word_template].presence
      @disposition = options.fetch(:disposition, 'attachment')
      @use_extras = options.fetch(:extras, false)
      @file_name = file_name(filename, options)
      @context = context
      define_template(filename, options)
      @content = options[:content] || @context.render_to_string(options)
      prerender_header_footer options
    end

    def send_file
      document = Htmltoword::Document.create(@content, @word_template, @use_extras, @header_footer)
      @context.send_data(document, filename: @file_name, type: Mime[:docx], disposition: @disposition)
    end

    private

    def prerender_header_footer(options)
      @header_footer = {}
      [:header, :footer].each do |hf_key|
        hf_options = options.fetch hf_key, {}
        next unless hf_options && hf_options[:html] && hf_options[:html][:template]

        hf_options[:html][:layout] ||= false
        render_opts = {
          :template => hf_options[:html][:template],
          :layout => hf_options[:html][:layout],
          :formats => hf_options[:html][:formats],
          :handlers => hf_options[:html][:handlers]
        }
        render_opts[:locals] = hf_options[:html][:locals] if hf_options[:html][:locals]
        @header_footer[hf_key] = @context.render_to_string(render_opts)
      end
    end

    def define_template(filename, options)
      if options[:template] == @context.action_name
        if filename =~ %r{^([^\/]+)/(.+)$}
          options[:prefixes] ||= []
          options[:prefixes].unshift $1
          options[:template] = $2
        else
          options[:template] = filename
        end
      end
    end

    def file_name(filename, options)
      name = options[:filename].presence || filename
      name =~ /\.docx$/ ? name : "#{name}.docx"
    end
  end
end
