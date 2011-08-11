'''
Created on Jul 12, 2011

@author: Tyler
'''

import os
import _winreg
import shutil

import xml.dom.minidom
import xml.dom
import xml.parsers.expat
import pickle

import old_setting_read

class XMLSettings( object ):
  
  
  
  XML_SETTING_VER = 1.2
  PREFS_LOC = None

  def __init__(self, file_path = None ):

    # Set location of xml file
    if file_path == None:
      if self.PREFS_LOC == None:
        raise Exception( 'No Config File specified' )
      else:
        self.cfg_file_path = self.PREFS_LOC
    else:
      self.cfg_file_path = file_path
      
    
    # Read in current state of xml file, if it exists
    try:
      print self.cfg_file_path
      self.cfgs = xml.dom.minidom.parse( self.cfg_file_path )
      print 'reading xml'
      
    except ( xml.parsers.expat.ExpatError, IOError ):
      impl = xml.dom.minidom.getDOMImplementation( )
      self.cfgs = impl.createDocument( None, 'settings', None )
      self.top_element = self.cfgs.documentElement
      self.top_element.setAttribute( 'version', str( self.XML_SETTING_VER ) )
      self.load_ver = str( self.XML_SETTING_VER )
      print 'new xml'
      
    else:
      self.top_element = self.cfgs.documentElement
      self.load_ver = self.top_element.getAttribute( 'version' )
      
    
    
    # If settings is old version, upgrade it
  
  def update_settings( self ):

    settings = { }
    for node in self.top_element.childNodes:
      if node.nodeType == xml.dom.Node.ELEMENT_NODE:
        
        setting_name = node.nodeName
        setting_val = old_setting_read.read_setting( self,
                                                     self.load_ver,
                                                     node )
        
        settings[ setting_name ] = setting_val
        
    while len( self.top_element.childNodes ) > 0:
      self.top_element.removeChild( self.top_element.childNodes[ 0 ] )  
        
    self.set_settings( settings )
    self.top_element.setAttribute( 'version', str( self.XML_SETTING_VER ) )
    self.write_settings( )
    
  
  def get_settings( self, settings ):
    """
    Returns:
      If a single valid setting name (str) is supplied, that setting's value
      If a list of setting names is supplied, a dict of setting values
      if no valid setting names are provided, returns None
    """
    data = None
    if isinstance( settings, ( list, tuple ) ):
      data = { }
      for setting in settings:
        setting = setting.replace( '_-_', ' ' )
        elements = self.cfgs.getElementsByTagName( setting )
        
        for element in elements:
          data[ setting ] = [ ]
          data[ setting ].append( self._get_setting( elements[ 0 ] ) )
          
          if data[ setting ] == [ ]:
            data[ setting ] = None
        
      if data == { }:
        return None
      
    if isinstance( settings, basestring ):
      settings = settings.replace( '_-_', ' ' )
      elements = self.cfgs.getElementsByTagName( settings )
      data = [ ]
      for element in elements:
        data.append( self._get_setting( element ) )

      if data == [ ]:
        return None
      
      if len( data ) == 1:
        return data[ 0 ]

    return data
  
  def _get_setting(self, element ):
    ''' 
    Get settings from this element, based off of the version
    '''
    # List Elements
    if element.getAttribute( 'type' ) == 'list':
      return_list = [ ]
      for node in element.getElementsByTagName( 'item' ):
        print 'an item'
        return_list.append( self._get_setting( node ) )
  
      return return_list
    
    # Dict Elements
    elif element.getAttribute( 'type' ) == 'dict':
      return_dict = { }
      for node in element.childNodes:
        if node.nodeType == node.ELEMENT_NODE:
          setting_name = node.tagName.replace( '_-_', ' ' ) 
          return_dict[ setting_name ] = self._get_setting( node )
  
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
    
    elif element.getAttribute( 'type' ) == 'int':
      if element.childNodes:
        if element.firstChild.nodeType == element.firstChild.TEXT_NODE:
          return int( element.firstChild.data )

      # This is most likely an error
      return None
        
    elif element.getAttribute( 'type' ) == 'float':
      if element.childNodes:
        if element.firstChild.nodeType == element.firstChild.TEXT_NODE:
          return float( element.firstChild.data )
        
      # Most likely an error
      return None
    
    # nonetype
    elif element.getAttribute( 'type' ) == 'none':
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
    
  def set_settings( self, settings = { }, append = False, write = False ):
    
    for setting in settings.keys( ):
      val = settings[ setting ]
      self._set_setting( setting, val, parent = None, append = append )
      
    if write == True:
      self.write_settings( )
    
  def write_settings( self ):

    if os.path.exists( self.cfg_file_path ):
      cfg_file_backup = shutil.copy2( self.cfg_file_path,
                                      '{0}.bak'.format( self.cfg_file_path ) )
    
    with open( self.cfg_file_path, 'wb' ) as cfg_file:
      self.cfgs.writexml( cfg_file)

    
  def _set_setting( self,
                    setting,
                    value,
                    parent = None,
                    append = False ):
    if parent == None:
      parent = self.top_element
    
    setting = setting.replace( ' ', '_-_')
    
    if parent.getElementsByTagName( setting ) and append == False:
      setting_node = parent.getElementsByTagName( setting )[ 0 ]
    else:
      setting_node = self.cfgs.createElement( setting )
      
    for child in setting_node.childNodes:
      setting_node.removeChild( child )
      
    self._create_setting( setting, setting_node, value, parent )
    
  def _create_setting( self, setting, setting_node, value, parent ):
    """
    Creating a separate create method so that it is very easy to extend in a
    child class.  _get_setting doesn't require this, as there's no preliminary
    setup like there is in _set_setting.
    
    Example:
    
    class 'Spam' wants to add an element for type 'Cheese'.
    Spam's def _create_setting would look like this:
    
    def _create_setting( self, setting, value, parent ):
      
      if isinstance( value, Cheese ):
        setting_node.setAttribute( 'type', 'cheese' )
        # Do things to set the node
      
      # Then run regular _parse_setting
      util.XMLSettings._parse_setting( self, setting, value, parent ) 
    """
    # List Values
    if isinstance( value, list ):
      setting_node.setAttribute( 'type', 'list' )
      for item in value:
        self._set_setting( 'item', item, parent = setting_node, append = True )
      parent.appendChild( setting_node )
        
    # Dict Values
    elif isinstance( value, dict ):
      setting_node.setAttribute( 'type', 'dict' )
      for k, v in value.iteritems( ):
        self._set_setting( k, v, setting_node )
      parent.appendChild( setting_node )
        
    # final leaf nodes
    else:
      # Bool Values
      if isinstance( value, bool ):
        setting_node.setAttribute( 'type', 'bool' )
      
      # Int values
      elif isinstance( value, int ):
        setting_node.setAttribute( 'type', 'int' )
        
      # Float values
      elif isinstance( value, float ):
        setting_node.setAttribute( 'type', 'float' )
      
      # string values   
      elif isinstance( value, basestring ):
        setting_node.setAttribute( 'type', 'string' )
      
      elif value == None:
        setting_node.setAttribute( 'type', 'none' )  
      
      else:
        raise Exception( 'type {0} is not currently a vaild type for storing settings'.format( type( value ) ) )
        
      setting_text = self.cfgs.createTextNode( str( value ) )
      setting_node.appendChild( setting_text )
      parent.appendChild( setting_node )

    
  def remove_settings_element( self, element ):
    
        pass