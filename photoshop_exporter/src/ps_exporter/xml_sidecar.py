'''
Created on Jul 17, 2011

@author: Tyler
'''

import win32com.client
import pythoncom
import os
import string
import random
import pickle
import re

import util

COMMENT_RE = '<!--(.*?)\|(.*?)-->'

def convert_rgba_to_num( rgba_str ):
  CONVERT_STR = 'rgba'
  
  for c in rgba_str:
    if c.lower( ) not in CONVERT_STR and c not in range( len ( CONVERT_STR ) ):
      return None
     
  out_str = rgba_str
  
  for i in range( len( CONVERT_STR ) ):
    out_str = out_str.replace( CONVERT_STR[ i ], str( i ) )
  return out_str
      
  

class XMLSidecar( util.XMLSettings ):
  """
  Interface for interacting with both XML settings for the sidecar file and the
  pickled string files used to store the map files themselves
  
  TODO: create a way to concatenate the map string files on to the end of the 
  sidecar XML file, probably in a comment
  """
  PREFS_LOC = None
  
  def __init__( self, file_path = None ):
    self._maps_to_write = { }
    self._maps = { }
    util.XMLSettings.__init__( self, file_path )
    if self.get_maps( ) == None:
      new_map = Map( )
      new_map.set_name( 'New Map' )
      self.add_map( new_map )
    
    print 'done loading'
        
    
  # Map functions
  def add_map( self, map ):
    """
    Add one map to the maps for the doc.
    """
    assert isinstance( map, Map ), 'map must be an instance of type Map'
    new_maps = self.get_maps( )
    if new_maps == None:
      new_maps = { }
    
    new_maps[ map.get_name( ) ] = ( map )
    self.set_maps( new_maps )
    
  def set_maps( self, maps ):
    """
    Set all of the maps to be used in this sidecar file
    """
    assert isinstance( maps, dict ), 'maps must be of type dict. type {0} was given.'.format( type( maps ))
    #self._maps = maps
    #self._set_map_settings_from_maps( )
    print 'maps: {0}'.format( maps )
    self.set_settings( { 'maps' : maps } )
    
  def remove_map( self, map ):
    """
    Remove a specific map from the sidecar file
    """
    maps = self.get_maps( )
    try:
      del maps[ map ]
    except KeyError:
      print 'That map does not exist'
    else:
      self.set_maps( maps )
      self._set_map_settings_from_maps( )
  
  def get_maps( self ):
    """
    return all of the maps in the sidecar file as an array
    """
    #return self._maps
    return self.get_settings( 'maps' )
  
  def _get_setting(self, element):
    if element.getAttribute( 'type' ) == 'map':
      if element.childNodes:
        settings = { }
        for node in element.childNodes:
          setting_name = node.tagName.replace( '_-_', ' ' ) 
          settings[ setting_name ] = self._get_setting( node )
          
        return_map = Map( settings[ 'export_file' ],
                          settings[ 'map_name' ] )
        
        return_map.set_action( settings[ 'action' ] ) 
        return_map.set_channels( settings[ 'channels' ] )
        return_map.set_resolution( settings[ 'resolution' ])
        
        self._maps[ settings[ 'map_name' ] ] = return_map
        
      else:
        # If node is empty
        return None

      return return_map

    elif element.getAttribute( 'type' ) == 'channelset':
      if element.childNodes:
        settings = { }
        for node in element.childNodes:
          setting_name = node.tagName.replace( '_-_', ' ' ) 
          settings[ setting_name ] = self._get_setting( node )
        return_cset = ChannelSet( settings[ 'shortname' ],
                                  settings[ 'chan_name' ],
                                  settings[ 'index' ],
                                  settings[ 'required' ] )
        
        return_cset.set_channels( settings[ 'channels' ] )
        
        return return_cset
        
      else:
        # Node is empty
        return None  
    
    elif element.getAttribute( 'type' ) == 'channel':
      if element.childNodes:
        settings = { }
        for node in element.childNodes:
          setting_name = node.tagName.replace( '_-_', ' ' ) 
          settings[ setting_name ] = self._get_setting( node )
       
        return_chan = Channel( settings[ 'shortname' ],
                               settings[ 'chan_name' ],
                               settings[ 'index' ],
                               settings[ 'required' ] )
        
        return_chan.set_layersets( settings[ 'layersets' ] )
        print 'break'
        return return_chan
          
      else:
        # Empty node
        return None
    
    else:
      # not a Map object
      return util.XMLSettings._get_setting(self, element)
  
  def _create_setting(self, setting, setting_node, value, parent):
   
    if isinstance( value, Map ):
      setting_node.setAttribute( 'type', 'map' )
      
      for k,v in value.get_settings( ).iteritems( ):

        self._set_setting( k, v, setting_node )

      parent.appendChild( setting_node )
      #self._maps_to_write[ value.get_name( ) ] = value
      
    elif isinstance( value, Channel ):
      setting_node.setAttribute( 'type', 'channel' )
      
      
      for k,v in value.get_settings( ).iteritems( ):
        self._set_setting( k, v, setting_node )

      parent.appendChild( setting_node )
    elif isinstance( value, ChannelSet ):
      setting_node.setAttribute( 'type', 'channelset' )
      
      for k,v in value.get_settings( ).iteritems( ):
        self._set_setting( k, v, setting_node )

      parent.appendChild( setting_node )

    else:  
      util.XMLSettings._create_setting(self, setting, setting_node, value, parent)
  
  def _map_rep_fn( self, match ):
    map = match.groups( )[ 0 ]
    if map in self._maps_to_write:
        data = self._maps_to_write[ map ]
        del self._maps_to_write[ map ]
    else:
        data = pickle.dumps( match.groups( )[ 1 ] )
   
  def write_settings(self):
    util.XMLSettings.write_settings( self )
    
class SidecarMap( util.XMLSettings ):
  PREFS_LOC = None
  pass
  
  
  

class Channel( object ):
  """
  
  """
  def __init__( self, shortname, name, index, required = True ):
    
    self._required = required
    self._index = index
    self._name = name
    self._shortname = shortname
    
    self._layersets = [ ]
    
  def get_settings( self ):
    """
    convenience function to use when converting to xml
    """
    settings_dict = { 'shortname' : self._shortname,
                      'chan_name' : self._name,
                      'required' : self._required,
                      'index' : self._index,
                      'layersets' : self._layersets }
    
    return settings_dict

  def get_shortname( self ):
    return self._shortname
  
  def get_name( self ):
    return self._name
  
  def add_layerset( self, layerset ):
    self._layersets.append( layerset, str( self._index ) )
      
  def remove_layerset( self, layerset ):
    for ls in self._layersets:
      if ls[ 0 ] == layerset:
    
  def get_layersets( self ):
    return self._layersets
  
  def set_layersets( self, layersets ):
    assert isinstance( layersets, dict )
    self._layersets = layersets
    
  def set_layer_channels( self, layerset, channels ):
    self._layersets[ layerset ] = str( channels )
    
  def get_layer_channels( self, layerset ):
    return self._layersets[ layerset ]
  
  def clear( self ):
    
    self._layersets = { }
    self._layer_channels = str( self.__index )
    
  def is_valid( self ):
    if not len( self._layersets ) == 0:
      return True
    
    return False
  
  def is_required( self ):
    return self._required
  
  def get_index( self ):
    return self._index

class ChannelSet( object ):
  
  def __init__( self, shortname, name, index, required = True ):
    
    self._shortname = shortname
    self._name = name
    self._required = required
    self._channels = { }
    self._index = index
    self._indicies = { }
    self._layer_channels = [ ]
    self._layersets = [ ]
    
    i = index
    for sn in self._shortname:
      self._channels[ sn ] = ( Channel( sn,
                                        self._name,
                                        i,
                                        self._required ) )
      self._indicies[ sn ] = i
      i = i + 1
  
  # TODO: set these up so they are more built around the channel set
  # so that channel set methods are not just relying on the child channels  
  
  def add_channel( self, channel ):
    self._channels.append( channel ) 
    
  def set_channels( self, channels ):
    self._channels = channels
  
  def get_channels( self ):
    return self._channels
  
  def get_channel( self, chan ):
    return self._channels[ chan ]
  
  def add_layerset( self, layerset ):
    for chan in self._channels.values( ):
      chan.add_layerset( layerset )
      
  def set_layersets( self, layersets ):
    for chan in self._channels.values( ):
      chan.set_layersets( layersets )
      
  def remove_layerset( self, layerset ):
    for chan in self._channels.values( ):
      chan.remove_layerset( layerset )
      
  def is_valid( self ):
    for chan in self._channels.values( ):
      if not chan.is_valid( ):
        return False
    return True
      
  def get_settings( self ):
    """
    convenience function to use when converting to xml
    """
    settings_dict = { 'shortname' : self._shortname,
                      'chan_name' : self._name,
                      'required' : self._required,
                      'channels' : self._channels,
                      'index' : self._index,
                      'indicies' : self._indicies,
                      'layer_channels' : self._layer_channels }
    
    return settings_dict
  
  
  
class Map( object ):
  """
  A class to make dealing with exporting maps from photoshop more convenient
  """
  
  # Valid Channels for the map
  __channels__ = [ [ 'r', 'red', True ],
                   [ 'g', 'green', True ],
                   [ 'b', 'blue', True ],
                   [ 'a', 'alpha', False ] ]

  
  def __init__( self, export_file = None, name = 'map' ):
    self._name = name
    self._export_file = export_file
    self._channels = { }
    self._action = None
    self._resolution = [ 512, 512 ]
    add = 0
    # Create channel attributes
    for i in range( len( self.__channels__) ):
      chan = self.__channels__[ i ]
      if len( chan[ 0 ] ) == 1:
        self._channels[ chan[ 0 ] ] = Channel( chan[ 0 ],
                                              chan[ 1 ],
                                              i + add,
                                              chan[ 2 ] )
      else:
        self._channels[ chan[ 0 ] ] = ChannelSet( chan[ 0 ],
                                                 chan[ 1 ],
                                                 i + add,
                                                 chan[ 2 ] )
        add = add + len( chan[ 0 ] ) - 1

  def get_settings( self ):
    """
    convenience function to use when converting to XML
    """
    settings_dict = { 'map_name' : self._name,
                      'channels' : self._channels,
                      'action' : self._action,
                      'resolution' : self._resolution,
                      'export_file' : self._export_file }
    
    return settings_dict
  
  # Name functions
  def get_name( self ):
    return self._name
  
  def set_name( self, name ):
    self._name = name
  
  #Export File
  def set_export_file( self, file ):
    self._export_file = file
  def get_export_file( self ):
    return self._export_file
  
  #PS Actions
  def set_action( self, action ):
    self.action = action  
  def remove_action( self ):
    self.action = None  
  def get_action( self ):
    return self.action 
  
  # Export Resolution
  def get_resolution( self ):
    return self._resolution
  
  def set_resolution( self, res ):
    self._resolution = [ int( res[ 0 ] ), int( res[ 1 ] ) ]

  
  def set_channels( self, channels ):
    """
    Really only have this here to ease loading from XML
    """
    self._channels = channels
  
  # Map Channel Assignment
  def set_channel_assignment( self, map_channels, layers, layer_channels ):   
    if isinstance( layers, basestring ):
      layers = [ layers ]
    if len( map_channels ) == 1:
      map_channel = map_channels[ 0 ]
      for layer in layers:
        self._channels[ map_channel ].add_layerset( layer )
        
      self._channels[ map_channel ].set_layer_channels( convert_rgba_to_num( layer_channels ) )

    else:
      assert len( map_channels ) == len( layer_channels)
      
      try:
        # See if we're dealing with a channel set
        self._channels[ map_channels ]
      except KeyError:
        # Assign layer channels to map channels
        for i in range( len( map_channels ) ):
          for layer in layers:
            self._channels[ map_channels[ i ] ].add_layerset( layer )
            
          self._channels[ map_channels[ i ] ].set_layer_channels( convert_rgba_to_num( layer_channels[ i ]))
    
      else:
        # For Channel set
        for layer in layers:
          self._channels[ map_channels ].add_layerset( layer )
          
        for i in range( len( map_channels ) ):
          chan = self._channels[ map_channels ].get_channel( map_channels[ i ] )
          chan.set_layer_channels( convert_rgba_to_num( layer_channels[ i ] ) )   
    return True
 
  # Channels    
  def get_channel( self, map_channel ):
    return self._channels[ map_channel ]  
  def get_channels( self ):
    return self._channels  
  def get_ordered_channels( self ):
    """
    Returns a list of channels, sorted by index number
    """
    channels = [ ]
    for channel in self._channels.values( ):
      while len( channels ) < channel.get_index( ) + 1:
        channels.append( None )    
      channels[ channel.get_index( ) ] = channel      
    return channels  
  def clear_channel( self, map_channel ):
    try:
      self._channels[ map_channel ].clear( )
    except KeyError:
      print 'map channel {0} does not exist'.format( map_channel )

  def is_valid( self ):
    """
    Returns True if all required channels are used, False if not.
    TODO: make this check whether the source layer sets/channels exist
    """
    for chan in self.channels:
      if chan.is_required( ) and not chan.is_valid( ):
        return False
      
    return True
  
  def export( self, file = None ):
    
    if file == None:
      export_file = self._export_file
    else:
      export_file = file
    
    ps_app = win32com.client.Dispatch( "Photoshop.Application")
    
    # Save settings
    save_units = ps_app.Preferences.RulerUnits
    save_dialogs = ps_app.DisplayDialogs
    
    # Set settings
    ps_app.Preferences.RulerUnits = 1
    ps_app.DisplayDialogs = 3
    
    # variables for accessing the docs
    active_doc = ps_app.ActiveDocument   
    export_doc = ps_app.Documents.Add( active_doc.Width, active_doc.Height )
    
    active_chans = active_doc.Channels
    active_layers = active_doc.Layers

    # So we can restore original visibility settings
    orig_layer_vis = [ ]
    orig_chan_vis = [ ]
    
    for layer in active_layers:
      orig_layer_vis.append( [ layer, layer.Visible ] )
    for chan in active_chans:
      orig_chan_vis.append( [ chan, chan.Visible ] )
    
    ps_app.ActiveDocument = active_doc
    
    # Hide all art layers
    for layer in active_doc.ArtLayers:
      layer.Visible = False
      
    # Create white background for copying 
    white_color = win32com.client.Dispatch( 'Photoshop.SolidColor.12' )
    white_color.RGB.HexValue = 'ffffff'
    
    # Fill the layer
    bg_layer = active_doc.ArtLayers.Add( )
    active_doc.ActiveLayer = bg_layer    
    active_doc.Selection.SelectAll( )
    active_doc.Selection.Fill( white_color, 2, 100 )
    
    # Move the layer to bottom
    
    try: active_doc.BackgroundLayer
    except pythoncom.com_error:
      # Place before
      place = 4
    else:
      # Place after
      place = 3
    bg_layer.Move( active_layers[ -1 ], place )
    
    # END OF SETUP
    
    for map_chan in self._channels.values():
      if isinstance( map_chan, Channel ):
        self.__copy_chan( map_chan, ps_app, active_doc, export_doc )
      elif isinstance( map_chan, ChannelSet ):
        for actual_chan in map_chan.get_channels( ).values( ):
          self.__copy_chan( actual_chan, ps_app, active_doc, export_doc )
          
      
          
    
    # Resize export canvas to desired size
    export_doc.ResizeImage( self.get_resolution( )[ 0 ],
                             self.get_resolution( )[ 1 ] )    
    
    # Export export doc
    tso = win32com.client.Dispatch( 'Photoshop.TargaSaveOptions' )
    chan_count = 0
    for chan in self._channels.values( ):
      if chan.is_valid( ):
        i = 1
        if isinstance( chan, ChannelSet ):
          i = len( chan.get_channels( ) )
        chan_count = chan_count + i
    
    assert chan_count <= 4 and chan_count >= 3, 'The exporter does not currently support maps with less than 3 or more than 4 channels' 
        
    if chan_count == 4:
      tso.Resolution = 32
    elif chan_count == 3:
      tso.Resolution = 24

    export_doc.SaveAs( export_file, tso, True, 2 )
    
    export_doc.Close( )
    
    # Delete the background white layer
    ps_app.ActiveDocument = active_doc
    bg_layer.Delete( )
    
    # Reset activedoc visibility to original
    for layer in orig_layer_vis:
      layer[ 0 ].Visible = layer[ 1 ]
      
    for chan in orig_chan_vis:
      chan[ 0 ].Visible = chan[ 1 ]
    
    # Reset settings
    ps_app.Preferences.RulerUnits = save_units
    ps_app.DisplayDialogs = save_dialogs
  
  def __copy_chan( self, map_chan, ps_app, source, dest ):
    export_sets = [ ]
    
    # Active Document must be active_doc for changing active doc layer/chan
    # Visibility
    ps_app.ActiveDocument = source     
       
    for set in source.LayerSets:
      if set.Name in map_chan.get_layersets( ):
        export_sets.append( set )
        set.Visible = True
        
      else:
        set.Visible = False

        
    if export_sets == [ ] and map_chan.is_required( ):
      print 'failed to find {0}'.format( map_chan.get_name( ) )
      dest.Close( )
      raise Exception( 'Error while exporting.  Required channel had no export sets.')
    
    elif not export_sets == [ ]:
    
      
      # make sure there are enough channels
      # Export doc must be active
      ps_app.ActiveDocument = dest
      while map_chan.get_index( ) > dest.Channels.Count - 1:
        # If the channel doesn't exist, add it
        dest.Channels.Add( )
        
      ps_app.ActiveDocument = source
      
      for i in range( source.Channels.Count ):

        if str( i ) not in map_chan.get_layer_channels( ):
          source.Channels[ i ].Visible = False
        else:
          source.Channels[ i ].Visible = True
          
      source.Selection.SelectAll( )
      source.Selection.Copy( True )
      
      ps_app.ActiveDocument = dest
      
      for i in range( dest.Channels.Count ):
        if not i == map_chan.get_index( ):
          dest.Channels[ i ].Visible = False
        else:
          dest.ActiveChannels = [ dest.Channels[ i ] ]
          dest.Channels[ i ].Visible = True
          
      dest.Paste( )
class MapAlt( dict ):
  """
  A subclass of dict that, when ready to use, will have either 3 or 4 elements.
  'r', 'g', and 'b' must be used for the map to be valid, 'a' can be used, but
  is not necessary.
  """
  
  # Valid Channels for the map
  __required_channels__ = 'rgb'
  __optional_channels__ = 'a'
  __channels__ = __required_channels__ + __optional_channels__
  
  def __setitem__( self, key, value ):
    except_string = 'Keys for maps must be one of {0}. You supplied {1}.'.format( ','.join( self.__channels__ ), key )
    assert key.lower( ) in self.__channels__, except_string
    
    dict.__setitem__( self, key.lower( ), value )
     
  def __getitem__( self, key ):
    except_string = 'Keys for maps must be one of {0}. You supplied {1}.'.format( ','.join( self.__channels__ ), key )
    assert key.lower( ) in self.__channels__, except_string
    
    dict.__getitem__( self, key )
  def is_valid( self ):
    """
    Returns True if all required channels are used, False if not
    """
    
    try:
      for channel in self.__required_channels__:
        self[ channel ]
    except KeyError:
      return False
    else:
      return True
    
  def export( self, location ):
    pass
    
class MapGreyscale( Map ):
  """
  A sublcass of Map that only has one channel.
  """
  __channels__ = [ [ 'r', 'red', True ] ]
  
class MapSimpleRGB( Map ):
  """
  A subclass of Map that has all 3 rgb channels grouped into a set
  """
  __channels__ = [ ['rgb', 'RGB', True ] ]
  
class MapSimpleRGBA( Map ):
  """
  A subclass of map that has 3 RGB channels grouped into a set, plus an alpha
  channel.
  """
  __channels__ = [ [ 'rgb', 'RGB', True ],
                   [ 'a', 'alpha', False ] ]
  
class MapRGB( Map ):
  """
  A subclass of map that has 3 RGB channels.
  """
  __channels__ = [ [ 'r', 'red', True ],
                   [ 'g', 'green', True ],
                   [ 'b', 'blue', True ] ]
  
