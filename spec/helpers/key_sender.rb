
require 'win32ole'

class KeySender
  def initialize(options={})
    @wsh = WIN32OLE.new('Wscript.Shell')
    @initial_wait = options[:initial_wait] || 0.2
  end

  # options:
  #     :timeout, :initial_wait, :if_target_missing, :silent, 
  def send(fenstername, tastenfolge, options={})
    initial_wait = options[:initial_wait] || @initial_wait
    timeout = options[:timeout] || 5
    start_time = Time.now
    sleep initial_wait

    ready_to_send = if fenstername
      loop do
        akt = @wsh.AppActivate(fenstername)
        break akt if akt
        break false if Time.now - start_time > timeout
        sleep 0.3
        print "-" unless options[:silent] 
      end
    else
      true # Keine Fenstername, immer senden'
    end

    if ready_to_send
      @wsh.SendKeys(tastenfolge) if tastenfolge
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

  def warte_auf_fenster(fensternamen, timeout=30)
    abbruch_zeit = Time.now + timeout
    loop do
      fensternamen.each do |fenstername|
        ready_to_send = @wsh.AppActivate(fenstername)
        return fenstername if ready_to_send
      end
      break false if Time.now > abbruch_zeit

      print " (noch #{'%.1f'%(abbruch_zeit - Time.now)}s) " unless options[:silent] 
      sleep 0.813
    end
  end
end
