
require 'win32ole'

class KeySender # :nodoc: #
  def initialize(window_name, options={})
    @window_name = window_name
    @wsh = WIN32OLE.new('Wscript.Shell')
    @initial_wait = options[:initial_wait] || 0.2
  end

  # options:
  #     :timeout, :initial_wait, :if_target_missing, :silent,
  def send(key_seq, options={})
    initial_wait = options[:initial_wait] || @initial_wait
    timeout = options[:timeout] || 5
    start_time = Time.now
    sleep initial_wait

    ready_to_send = if @window_name
      loop do
        akt = @wsh.AppActivate(@window_name)
        break akt if akt
        break false if Time.now - start_time > timeout
        sleep 0.3
        print "-" unless options[:silent]
      end
    else
      true # Keine Window_name, immer senden'
    end

    if ready_to_send
      @wsh.SendKeys(key_seq) if key_seq
      yield if block_given?
    else
      else_aktion = options[:if_target_missing]
      case else_aktion
      when Proc
        else_aktion.call
      when nil
      else
        raise else_aktion
      end
    end
    ready_to_send
  end

  def wait_for_window(windowname, timeout=30)
    break_time = Time.now + timeout
    loop do
      windowname.each do |window_name|
        ready_to_send = @wsh.AppActivate(window_name)
        return window_name if ready_to_send
      end
      break false if Time.now > break_time

      print " (noch #{'%.1f'%(break_time - Time.now)}s) " unless options[:silent]
      sleep 0.813
    end
  end
end

if __FILE__ == $0
  key_sender = KeySender.new(ARGV[1])
  while not $stdin.eof? do
    key_sequence = $stdin.gets.chomp
    key_sender.send key_sequence
  end
end
