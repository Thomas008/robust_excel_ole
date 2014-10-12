require 'tmpdir'

def create_tmpdir    
  tmpdir = Dir.mktmpdir
  FileUtils.cp_r(File.join(File.dirname(__FILE__), '../data'), tmpdir)
  tmpdir + '/data/'
end

def rm_tmp(tmpdir)    
  FileUtils.remove_entry_secure(File.dirname(tmpdir))
end
