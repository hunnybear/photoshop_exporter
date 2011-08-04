'''
Created on Jul 14, 2011

@author: Tyler
'''

import pickle

def read_setting( self, version, element ):
  ''' 
  Get settings from this element, based off of the version
  '''
  if version == '1.0.0':
    
    # List Elements
    if element.getAttribute( 'type' ) == 'list':
      return_list = [ ]
      for node in element.getElementsByTagName( 'item' ):
        return_list.append( read_setting( self, version, node ) )
  
      return return_list
    
    # Dict Elements
    elif element.getAttribute( 'type' ) == 'dict':
      return_dict = { }
      for node in element.childNodes:
        if node.nodeType == node.ELEMENT_NODE:
  
          return_dict[ node.tagName ] = read_setting( self, version, node )
  
      return return_dict
    
    elif element.getAttribute( 'type' ) == 'bool':
      # First see if the element has any data
      if element.childNodes:
        if element.firstChild.nodeType == element.firstChild.TEXT_NODE:
          data = element.firstChild.data 
        else:
          # This is most likely an error
          data = None
      else:
        # If the node is empty
        data = None
      if data == 'True':
        return True
      elif data == 'False':
        return False
      else:
        return None
      
  
    # String Elements
    else:
      # First see if the element has any data
      if element.childNodes:
        if element.firstChild.nodeType == element.firstChild.TEXT_NODE:
          data = element.firstChild.data 
        else:
          # This is most likely an error
          data = None
      else:
        # If the node is empty
        data = ''
      return data
  if version == '1.1':
    # First, see if element has any data
    if element.childNodes:
      if element.firstChild.nodeType == element.firstChild.TEXT_NODE:
        pickled_data = element.firstChild.data
        
        # Unpickle and return the data
        return pickle.loads( pickled_data )
      else:
        # This is most likely an error when this setting was written
        return None
      
    else:
      return None