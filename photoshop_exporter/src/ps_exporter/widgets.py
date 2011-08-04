'''
Created on Jul 24, 2011

@author: Tyler
'''
import Tkinter
import win32com.client
import functools

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
    
    self.sets_frame = Tkinter.Frame( self.border_frame, padx = 10 )
    self.sets_frame.grid( row = 0, column = 2 )
    
    Tkinter.Frame( self.border_frame,
                   width = 2,
                   borderwidth = 2,
                   relief = Tkinter.GROOVE ).grid( row = 0,
                                                   column = 1,
                                                   sticky = Tkinter.N + Tkinter.S)
    
    self.chan_frame = Tkinter.Frame( self.border_frame, padx = 10 )
    self.chan_frame.grid( row = 0, column = 4)
    
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
    
    chan_row = 0
    chan_col = 0
    
    #==========================================================================
    # Column 2 - Layerset
    #==========================================================================
    doc_layersets = [ ]
    for ls in self._app.ps.ActiveDocument.LayerSets:
      doc_layersets.append( ls.Name )
      
    self._chan_layersets = self._channel.get_layersets( )
    ls_widgets = [ ]
    self._ls_vars = [ ]
    
    ls_i = 0
    for ls in self._channel.get_layersets( ):
      var = Tkinter.StringVar( self._app ) 
      
      var.set( ls )
      print ls_i
      ls_om = Tkinter.OptionMenu( self.sets_frame,
                                  var,
                                  *doc_layersets,
                                  command = functools.partial( self.handle_ls_change,
                                                               ls_i  ) )
      ls_om.config( width = 30 )
      ls_om.grid( row = ls_i, column = 0, rowspan = 2 )
      ls_widgets.append( ls_om )
      self._ls_vars.append( var )
      ls_i = ls_i + 2
      
      #==========================================================================
      # Column 3 - Channels
      #==========================================================================
  
      doc_channels = self._app._export_doc.Channels
      
      
      
      for channel in doc_channels:
        var = Tkinter.IntVar( )
        cbx = Tkinter.Checkbutton( self.chan_frame,
                                   text = channel.Name,
                                   variable = var )
        cbx.grid( row = chan_row, column = chan_col, sticky = Tkinter.W )
        
        self._chan_cbxs.append( cbx )
        self._chan_vars.append( var )
        
        if chan_row == 0:
          chan_row = 1
        else:
          chan_row = 0
          chan_col = chan_col + 1
      
  def handle_add_layerset( self ):
    """
    Handle Callback for the 'add layerset' button
    """
    self._channel.add_layerset( None )

    self.refresh( )
    
  def handle_ls_change( self, item, name ):
    """
    Handle changing a layerset in the channel to another
    """

    self._chan_layersets[ item ] = name
    self._channel.set_layersets( self._chan_layersets )

      
  
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
      resx = self._app.ps.ActiveDocument.Width
      resy = self._app.ps.ActiveDocument.Height
      
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
    
    
    