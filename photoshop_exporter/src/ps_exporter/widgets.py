'''
Created on Jul 24, 2011

@author: Tyler
'''
import Tkinter
import win32com.client
import functools
import timeit

class Widget( Tkinter.Frame ):
  """
  Abstract base class used as base for widgets
  """
  
  def __init__( self, parent, app, *args, **kwargs ):
    
    self._parent = parent
    self._app = app
    Tkinter.Frame.__init__( self, parent, *args, **kwargs )
    
    self.build( )
    
  def build( self ):
    """
    Placeholder method.  Used in non-abstract classes to build the widget UI.
    """
    
class ChannelWidget( Widget ):
  """
  UI widget to edit a channel of a map.
  """
  
  def __init__( self, parent, app, channel ):
    self._channel = channel
    self._set_widgets = [ ]
    self._chan_cbxs = [ ]
    self._chan_vars = [ ]
    self.parent = parent
    print channel.get_layersets()
    Widget.__init__( self, parent, app, pady = 3 )
    
  def add_layerset( self ):
    """
    add a new layerset to the ChannelWidget
    """
    pass
    
  def build( self ):
    """
    Build the UI
    """
    
    self.border_frame = Tkinter.Frame( self,
                                       relief = Tkinter.RAISED,
                                       borderwidth = 2)
    self.border_frame.grid( row = 0,
                            column = 0,
                            sticky = Tkinter.N + Tkinter.E + Tkinter.S + Tkinter.W )
    
    self.name_frame = Tkinter.Frame( self.border_frame, padx = 10, pady = 5 )
    self.name_frame.grid( row = 0, column = 0 )
    
    Tkinter.Frame( self.border_frame,
                   width = 2,
                   borderwidth = 2,
                   relief = Tkinter.GROOVE ).grid( row = 0,
                                                   column = 1,
                                                   sticky = Tkinter.N + Tkinter.S)
    
    self.sets_frame = Tkinter.Frame( self.border_frame )
    self.sets_frame.grid( row = 0, column = 2 )
    
    
    
    #==========================================================================
    # Column 1 - Name
    #==========================================================================
    
    name = Tkinter.Label( self.name_frame, text = self._channel.get_name( ) )
    name.grid( row = 0, column = 0 )
    
    btn_add_layerset = Tkinter.Button( self.name_frame,
                                       text = 'Add Group',
                                       command = self.handle_add_layerset )
    btn_add_layerset.grid( row = 1, column = 0 )
    
      
    
    

    if len( self._channel.get_layersets( ) ) == 0:
      self._channel.add_layerset( None )
      
    self.refresh( )
    
  def refresh( self ):
    
    
    
    #==========================================================================
    # Column 2 - Layerset
    #==========================================================================
    doc_layersets = [ ]
    used_layersets = list( self._channel.get_layersets( ) )
    for ls in self._app.ps.ActiveDocument.LayerSets:
      doc_layersets.append( ls.Name )
    ls_widgets = [ ]
    self._ls_vars = [ ]
    
    ls_i = 0
    for ls in used_layersets:
      if not ls == None:
        used = set( used_layersets ) - set( ls )
      else:
        used = set( used_layersets )
      om_layersets = [ x for x in doc_layersets if x not in used ]
      
      ls_frame = Tkinter.Frame( self.sets_frame )
      ls_frame.grid( column = 0, row = ls_i )
      var = Tkinter.StringVar( self._app ) 
      var.set( ls )
      ls_om = Tkinter.OptionMenu( ls_frame,
                                  var,
                                  *om_layersets,
                                  command = functools.partial( self.handle_ls_change,
                                                               ls  ) )
      ls_om.config( width = 30 )
      
      ls_widgets.append( ls_om )
      self._ls_vars.append( var )
      ls_i = ls_i + 1
      
      ls_om.grid( row = 0, column = 0 )

      # Divider
      Tkinter.Frame( ls_frame,
                     width = 2,
                     borderwidth = 2,
                     relief = Tkinter.GROOVE ).grid( row = 0,
                                                     rowspan = 1,
                                                     column = 1,
                                                     sticky = Tkinter.N + Tkinter.S)
      
      #==========================================================================
      # Column 3 - Channels
      #==========================================================================

      
      chan_cbx = MapChannelCheckBoxWidget( ls_frame, self._app, self._channel, ls )
      chan_cbx.grid( row = 0, column = 2 )

          
      
                     
      
          
   

      
  def handle_add_layerset( self ):
    """
    Handle Callback for the 'add layerset' button
    """
    self._channel.add_layerset( None )

    self.refresh( )
    
  def handle_ls_change( self, orig, name ):
    """
    Handle changing a layerset in the channel to another
    """

    new_ls = self._channel.get_layersets( )
    new_ls[ name ] = new_ls[ orig ]
    del new_ls[ orig ]
    self._channel.set_layersets( new_ls )
    self.refresh( )

class MapChannelCheckBoxWidget( Widget ):
  
  def __init__( self, parent, app, channel, layerset, *args, **kwargs ):
    
    self._map_channel = channel
    self._boxes = [ ]
    self._layerset = layerset
    Widget.__init__( self, parent, app, *args, **kwargs )
  
  def build( self ):
    doc_channels = self._app._active_doc.Channels
    self._cbx_vars = [ ]
    
    row = 0
    col = 0
    for i in range( doc_channels.Count ):
      chan = doc_channels[ i ]
      var = Tkinter.IntVar( )
      self._cbx_vars.append( var )
      cbx = Tkinter.Checkbutton( self,
                                 text = chan.Name,
                                 variable = var,
                                 command = self._handle_cbx )
      
      cbx.grid( row = row, column = col, sticky = Tkinter.W )
      
      if col == 0:
        col = 1
      else:
        col = 0,
        row = row + 1
    self.refresh( )
        
  def refresh( self ):
    for i in range( len ( self._cbx_vars ) ):
      if str( i ) in self._map_channel.get_layer_channels( self._layerset ):
        self._cbx_vars[ i ] = True
      else:
        self._cbx_vars[ i ] = False
    
      
  def _handle_cbx( self ):
    for i in range( len ( self._cbx_vars ) ):
      chan_str = ''
      if self._cbx_vars[ i ].get( ) == True:
        chan_str = chan_str + str( i )
    self._map_channel.set_layer_channels( self._layerset, chan_str )
  
class MapInfoWidget( Widget ):
  """
  UI widget used to edit the map info
  """
  
  def __init__( self, parent, app, map, *args, **kwargs ):
    
    self._map = map
    
    Widget.__init__( self, parent, app, *args, **kwargs )
  
  def build( self ):
    
    self.col0 = Tkinter.Frame( self, padx = 5 )
    self.col0.grid( row = 0, column = 0 )
    
    self.col1 = Tkinter.Frame( self, padx = 5 )
    self.col1.grid( row = 0, column = 1 )
    
    self.col2 = Tkinter.Frame( self, padx = 5 )
    self.col2.grid( row = 0, column = 2 )
    
    #==========================================================================
    # Column 0
    #==========================================================================
    
    self.tbx_name = Tkinter.Text( self.col0, width = 35, height = 1 )
    self.tbx_name.grid( row = 0, column = 0)
    
    self.btn_del_map = Tkinter.Button( self.col0, text = 'Remove Map' )
    self.btn_del_map.grid( row = 1, column = 0, sticky  = Tkinter.E )
  
    #==========================================================================
    # Column 1
    #==========================================================================
    var = Tkinter.StringVar( )
    options = [ 'rgb', 'rgba' ]
    self.om_type = Tkinter.OptionMenu( self.col1, var, *options )
    self.om_type.config( width = 10 )
    self.om_type.grid( row = 0, column = 0, columnspan = 3 )
    
    if None in self._map.get_resolution( ):
      resx = int( self._app.ps.ActiveDocument.Width )
      resy = int( self._app.ps.ActiveDocument.Height )
      self._map.set_resolution( [ resx, resy ] )

      
    else:
      resx = self._map.get_resolution( )[ 0 ]
      resy = self._map.get_resolution( )[ 1 ]
      
    
    self.tbx_width = Tkinter.Text( self.col1, width = 5, height = 1 )
    self.tbx_width.insert( Tkinter.END, resx )
    self.tbx_width.grid( row = 1, column = 0, sticky = Tkinter.E )
    Tkinter.Label( self.col1, text = 'x' ).grid( row = 1, column = 1 )
    self.tbx_height = Tkinter.Text( self.col1, width = 5, height = 1 )
    self.tbx_height.insert( Tkinter.END, resy )
    self.tbx_height.grid( row = 1, column = 2, sticky = Tkinter.W )
    
class MapWidget( Widget ):
  """
  The main Map Info Widget
  """
  
  def __init__( self, parent, app, map ):
    self._map = map
    self.channel_widgets = [ ]
    
    Widget.__init__( self,
                     parent,
                     app,
                     relief = 'raised',
                     padx = 10,
                     pady = 5,
                     borderwidth = 2 )
    

  def build( self ):
    map_info = MapInfoWidget( self, self._app, self._map, padx = 5, pady = 5 )
    map_info.grid( row = 0, column = 0 )
    
    pad = Tkinter.Frame( self,
                         pady = 2 )
    
    pad.grid( row = 1,
              column = 0,
              sticky = Tkinter.E + Tkinter.W )
    
    Tkinter.Frame( pad,
                   width = 2,
                   borderwidth = 2,
                   relief = Tkinter.GROOVE ).grid( row = 1,
                                                   column = 0,
                                                   sticky = Tkinter.E + Tkinter.W)
    
    
    self.refresh( )

  def refresh( self ):
    print self._map
    for channel_widget in self.channel_widgets:
      channel_widget.destroy( )

    row =2
    for channel in self._map.get_ordered_channels( ):
      cw = ChannelWidget( self, self._app, channel )
      cw.grid( row = row, column = 0 )
      self.channel_widgets.append( cw )
      row = row + 1
      
  
  
  def get_map( self ):
    return self._map
    
    
    