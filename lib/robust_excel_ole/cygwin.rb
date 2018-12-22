module RobustExcelOle
  module Cygwin
    require 'Win32API'

    @conv_to_full_posix_path =
      Win32API.new('cygwin1.dll', 'cygwin_conv_to_full_posix_path', 'PP', 'I')
    @conv_to_posix_path =
      Win32API.new('cygwin1.dll', 'cygwin_conv_to_posix_path', 'PP', 'I')
    @conv_to_full_win32_path =
      Win32API.new('cygwin1.dll', 'cygwin_conv_to_full_win32_path', 'PP', 'I')
    @conv_to_win32_path =
      Win32API.new('cygwin1.dll', 'cygwin_conv_to_win32_path', 'PP', 'I')

    def cygpath(options, path) # :nodoc:
      absolute = shortname = false
      func = nil
      options.delete(" \t-").chars do |opt|
        case opt
        when 'u'
          func = [@conv_to_full_posix_path, @conv_to_posix_path]
        when 'w'
          func = [@conv_to_full_win32_path, @conv_to_win32_path]
        when 'a'
          absolute = true
        when 's'
          shortname = true
        end
      end
      raise 'first argument must contain -u or -w' if func.nil?

      func = absolute ? func[0] : func[1]
      buf = "\0" * 300
      raise 'cannot convert path name' if func.Call(path, buf) == -1

      buf.delete!("\0")
      buf
    end
    module_function :cygpath
  end
end
