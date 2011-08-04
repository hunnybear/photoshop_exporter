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
    
    if not self._get_map_settings( ) == None:
      for map in self._get_map_settings( ):
        print map
        map_file = '{0}.{1}.map'.format( self.cfg_file_path,
                                         map )
        with open( map_file, 'rb' ) as f:
          self._maps[ map ] = pickle.load( f )
    else:
      new_map = Map( )
      new_map.set_name( 'New Map' )
      self.add_map( new_map )
        
    
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
    assert isinstance( maps, dict ), 'maps must be of type dict'
    self._maps = maps
    self._set_map_settings_from_maps( )
    
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
      
  def _set_map_settings_from_maps( self ):
    self.set_settings( {'maps' : self.get_maps( ) } ) 

  def _get_map_settings( self ):
    """
    return all of the names of maps listed in the xml sidecar
    """
    return self.get_settings( 'maps' )
  
  def get_maps( self ):
    """
    return all of the maps in the sidecar file as an array
    """
    return self._maps
  
  def _get_setting(self, element):
    
    if element.getAttribute( 'type' ) == 'map':
      if element.childNodes:
        data = element.firstChild.data
      else:
        # If node is empty
        print 'empty node'
        return None
      
      map_file = '{0}.{1}.map'.format( self.cfg_file_path,
                                       data )
      with open( map_file, 'rb' ) as f:
        return pickle.load( f )
      
      #with open( self.cfg_file_path, 'rb' ) as f:
      #  file_text = f.read( )
        
        #map_comments = re.compile( COMMENT_RE, re.DOTALL )
        #matches = map_comments.findall( file_text )
        #print matches
        #if not matches == None:
        #  for match in matches: 
        #    if match[ 0 ] == data:
        #      print match[ 1 ]
        #      return pickle.loads( match[ 1 ] ) 
      
        #else:
        #  print 'didnt match setting'
        #  return None
    else:
      # not a Map object
      return util.XMLSettings._get_setting(self, element)
  
  def _create_setting(self, setting, setting_node, value, parent):
   
    if isinstance( value, Map ):
      setting_node.setAttribute( 'type', 'map' )
      
      text_node = self.cfgs.createTextNode( value.get_name( ) )
      setting_node.appendChild( text_node )
      parent.appendChild( setting_node )
      self._maps_to_write[ value.get_name( ) ] = value
    
    else:  
      util.XMLSettings._create_setting(self, setting, setting_node, value, parent)
  
  def _map_rep_fn( self, match ):
    map = match.groups( )[ 0 ]
    if map in self._maps_to_write:
        data = self._maps_to_write[ map ]
        del self._maps_to_write[ map ]
    else:
        data = pickle.dumps( match.groups( )[ 1 ] )
    return '<!--{0}|{1}-->'.format( map, data )
    
  def write_settings(self):
    util.XMLSettings.write_settings( self )
    
    # Attempt to append pickled string at end of file.  Not successful so far
    #with open( self.cfg_file_path, 'rb' ) as f:
    #  file_text = f.read( )
      
      #map_comments = re.compile( COMMENT_RE, re.DOTALL )
      #if not map_comments.match( file_text ) == None:
      #  file_text = map_comments.sub( self._map_rep_fn( file_text ) )

      #new_string = ''
      #for k,v in self._maps_to_write.iteritems( ):
      #  new_string = '{0}<!--{1}|{2}-->\n'.format( new_string,
      #                                             k,
      #                                             pickle.dumps( v ) )
      
      #file_text = '{0}\n{1}'.format( file_text, new_string )
      
    #with open( self.cfg_file_path, 'wb' ) as f:
    #  f.write( file_text )
    
    for k, v in self._maps_to_write.iteritems( ):
      filename = '{0}.{1}.map'.format( self.cfg_file_path,
                                        k )
      f = open( filename, 'wb' )
      pickle.dump( v, f )

    
class SidecarMap( util.XMLSettings ):
  PREFS_LOC = None
  pass
  
  
  

class Channel( object ):
  """
  
  """
  def __init__( self, shortname, name, index, required = True ):
    
    self.__required = required
    self.__index = index
    self.__name = name
    self.__shortname = shortname
    
    self.__layersets = [ ]
    
    self.__layer_channels = str( index )

  def get_shortname( self ):
    return self.__shortname
  
  def get_name( self ):
    return self.__name
  
  def add_layerset( self, layerset ):
    self.__layersets.append( layerset )
      
  def remove_layerset( self, layerset ):
    self.__layersets.remove( layerset )
    
  def get_layersets( self ):
    return self.__layersets
  
  def set_layersets( self, layersets ):
    self.__layersets = layersets
    
  def set_layer_channels( self, channels ):
    self.__layer_channels = channels
    
  def get_layer_channels( self ):
    return self.__layer_channels
  
  def clear( self ):
    
    self.__layersets = [ ]
    self.__layer_channels = str( self.__index )
    
  def is_valid( self ):
    if not len( self.__layers ) == 0:
      return True
    
    return False
  
  def is_required( self ):
    return self.__required
  
  def get_index( self ):
    return self.__index

class ChannelSet( object ):
  
  def __init__( self, shortname, name, index, required = True ):
    
    self._shortname = shortname
    self._name = name
    self._required = required
    self._channels = { }
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
      
  def remove_layerset( self, layerset ):
    for chan in self._channels.values( ):
      chan.remove_layerset( layerset )

  
  
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
    self.channels = { }
    self.action = None
    self.resolution = [ None, None ]
    
    add = 0
    # Create channel attributes
    for i in range( len( self.__channels__) ):
      chan = self.__channels__[ i ]
      if len( chan[ 0 ] ) == 1:
        self.channels[ chan[ 0 ] ] = Channel( chan[ 0 ],
                                              chan[ 1 ],
                                              i + add,
                                              chan[ 2 ] )
      else:
        self.channels[ chan[ 0 ] ] = ChannelSet( chan[ 0 ],
                                                 chan[ 1 ],
                                                 i + add,
                                                 chan[ 2 ] )
        add = add + len( chan[ 0 ] ) - 1
    
    
  
    print( dir( self ) )
  
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
  def add_action( self, action ):
    self.action = action  
  def remove_action( self ):
    self.action = None  
  def get_action( self ):
    return self.action 
  
  # Export Resolution
  def get_resolution( self ):
    return self.resolution
  def set_resolution( self, resolution ):
    self.resolution = resolution 
  
  # Map Channel Assignment
  def set_channels( self, map_channels, layers, layer_channels ):   
    if isinstance( layers, basestring ):
      layers = [ layers ]
    if len( map_channels ) == 1:
      map_channel = map_channels[ 0 ]
      for layer in layers:
        self.channels[ map_channel ].add_layerset( layer )
        
      self.channels[ map_channel ].set_layer_channels( convert_rgba_to_num( layer_channels ) )

    else:
      assert len( map_channels ) == len( layer_channels)
      
      try:
        self.channels[ map_channels ]
      except KeyError:
        # Assign layer channels to map channels
        for i in range( len( map_channels ) ):
          for layer in layers:
            self.channels[ map_channels[ i ] ].add_layerset( layer )
            
          self.channels[ map_channels[ i ] ].set_layer_channels( convert_rgba_to_num( layer_channels[ i ]))
    
      else:
        for layer in layers:
          self.channels[ map_channels ].add_layerset( layer )
          
        for i in range( len( map_channels ) ):
          chan = self.channels[ map_channels ].get_channel( map_channels[ i ] )
          chan.set_layer_channels( convert_rgba_to_num( layer_channels[ i ] ) )   
    return True
 
  # Channels    
  def get_channel( self, map_channel ):
    return self.channels[ map_channel ]  
  def get_channels( self ):
    return self.channels  
  def get_ordered_channels( self ):
    """
    Returns a list of channels, sorted by index number
    """
    channels = [ ]
    for channel in self.channels.values( ):
      while len( channels ) < channel.get_index( ) + 1:
        channels.append( None )    
      channels[ channel.get_index( ) ] = channel      
    return channels  
  def clear_channel( self, map_channel ):
    try:
      self.channels[ map_channel ].clear( )
    except KeyError:
      print '{0} is not a valid channel'.format( map_channel )
  
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
    
    ps_app = win32com.client.Dispatch( "Photoshop.Application")
    
    # Save units settings
    save_units = ps_app.Preferences.RulerUnits
    
    # Set ruler units to pixels
    ps_app.Preferences.RulerUnits = 1
    
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
    print type( place )
    print active_layers[ -1 ].Name
    bg_layer.Move( active_layers[ -1 ], place )
    
    # END OF SETUP
    
    for map_chan in self.channels.values():
      if isinstance( map_chan, Channel ):
        self.__copy_chan( map_chan, ps_app, active_doc, export_doc )
      elif isinstance( map_chan, ChannelSet ):
        for actual_chan in map_chan.get_channels( ).values( ):
          self.__copy_chan( actual_chan, ps_app, active_doc, export_doc )
          
      
          
    
    # Resize export canvas to desired size
    export_doc.ResizeImage( self.resolution[ 0 ], self.resolution[ 1 ] )    
    
    # Export export doc
    
    # Delete the background white layer
    ps_app.ActiveDocument = active_doc
    bg_layer.Delete( )
    
    # Reset activedoc visibility to original
    for layer in orig_layer_vis:
      layer[ 0 ].Visible = layer[ 1 ]
      
    for chan in orig_chan_vis:
      chan[ 0 ].Visible = chan[ 1 ]
    
    ps_app.Preferences.RulerUnits = save_units 
    #export_doc.Close( )
  
  def __copy_chan( self, map_chan, ps_app, source, dest ):
    print 'copying {0}'.format(  map_chan.get_name( ) )
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
      return None
    
    elif not export_sets == [ ]:
    
      
      # make sure there are enough channels
      # Export doc must be active
      ps_app.ActiveDocument = dest
      while map_chan.get_index( ) > dest.Channels.Count - 1:
        # If the channel doesn't exist, add it
        dest.Channels.Add( )
        
      ps_app.ActiveDocument = source
      
      for i in range( source.Channels.Count ):
        
        print 'setting chan vis'
        if str( i ) not in map_chan.get_layer_channels( ):
          source.Channels[ i ].Visible = False
        else:
          source.Channels[ i ].Visible = True
          
      source.Selection.SelectAll( )
      source.Selection.Copy( True )
      
      ps_app.ActiveDocument = dest
      
      for i in range( dest.Channels.Count ):
        print dest.Channels[ i ].Name
        if not i == map_chan.get_index( ):
          print '{0} not visible'.format( i )
          dest.Channels[ i ].Visible = False
        else:
          print '{0} visible'.format( i )
          
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

