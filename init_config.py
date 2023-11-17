
# init_config.py -- set up configuration files

import dotenv
from types import ModuleType

def init_config(config, env_file_name=".env", env_required=True):
    """ initialize configuration system.  Return a dict-like object

        config - a dict like object; one alternative is to use an imported config.py object.  Keys starting with '__' will be ignored
        env_file_name - a dotenv initialization file; by default named ".env".  Values in this file will overwrite values in config
        env_required - if true (the default) it will be an error is the env_file_name is not found

    """

   
    class AttrDict(dict):
        """ A dictionary that also allows access via dict.key in addition to dict[key]
        """
        def __init__(self, *args, **kwargs):
            super(AttrDict, self).__init__(*args, **kwargs)
            self.__dict__ = self

    config_dotenv = dotenv.dotenv_values(env_file_name, verbose=env_required)

    result = AttrDict()

    #print(f"instance config { isinstance(config, dict) } { type(config) }")
    #print(f"instance config_dotenv { isinstance(config_dotenv, dict) } { type(config_dotenv) }")
    #print(f"instance result.__module__ { isinstance(result.__module__, dict) } { type(result.__module__) }")

    config_append(result, config)
    config_append(result, config_dotenv)

    return result



def config_append(result, new):

    if new is None:
        return

    if isinstance(new, dict):
        for key, val in new.items():
            if not key.startswith('__'):
                for key, val in new.items():
                    result[key] = val

    elif isinstance(new, ModuleType):
        for item in dir(new):
            if not item.startswith('__'):
                #print(f"init_config: item { item }")
                result[item] = getattr(new, item)

    else:
        raise Exception("unknown type passed: { type(new) }: not a dict or module")

    return result

