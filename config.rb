require 'local_config'

# gets default settings overwritten with local config hash
def get_config
  {}.update(local_config)
end