'''
Created on Jul 7, 2011

@author: Tyler
'''
import Tkinter
import tkMessageBox
import win32com.client
import pythoncom
import pywintypes
import os

import widgets
import xml_sidecar

SIDECAR_EXT = 'sde'

class Export( object ):
  """
  Check if export settings exist for the current PS file.  If it does exist,
  export files, if it doesn't, open the Export Settings UI
  """
  
  def __init__( self ):
    self._ps = win32com.client.Dispatch( 'Photoshop.Application' )
    try:
      self._active_doc = self._ps.ActiveDocument
    except pythoncom.com_error:
      root = Tkinter.Tk()
      root.withdraw() 
      tkMessageBox.showerror( 'No file open',
                              'There is currently no file open in Photoshop.',
                              parent = None )
      
      raise Exception( 'No file currently open in Photoshop' )
      
    self._doc_loc = self._active_doc.FullName
    self._sidecar_loc = '{0}.{1}'.format( self._doc_loc, SIDECAR_EXT )
    if os.path.exists( self._sidecar_loc ):
      self.do_export( )
    else:
      export_settings = ExportSettingsUI( doc = self._active_doc )
      
      
  def do_export( self ):
    """
    Export the doc.
    TODO: put up a dialog to ask which maps to export. Make this so user can 
    choose to not see this dialog
    """
    
    self.sidecar = xml_sidecar.XMLSidecar( self._sidecar_loc )
    print 'doing export'
    for map in self.sidecar.get_maps( ).values( ):
      map.export( )
  

class ExportSettingsUI( Tkinter.Frame ):
  
  def __init__( self, master = None, doc = None ):

    Tkinter.Frame.__init__( self, master )
    
    self._maps = { }
    self.ps = win32com.client.Dispatch( 'Photoshop.Application' )
    
    # Set up doc and sidecar
    if doc == None:
      try: self._active_doc = self.ps.ActiveDocument
      except pythoncom.com_error:

        root = Tkinter.Tk()
        root.withdraw() 
        tkMessageBox.showerror( 'No file open',
                                'There is currently no file open in Photoshop.',
                                parent = None )
      
        raise Exception( 'No file currently open in Photoshop' )
    else:
      self._active_doc = doc
    
    self._sidecar_loc = '{0}.{1}'.format( self._active_doc.FullName,
                                          SIDECAR_EXT )  
    
    self._sidecarXML = xml_sidecar.XMLSidecar( self._sidecar_loc )
    
    self._doc_loc = self._active_doc.FullName
    self._sidecar_loc = '{0}.{1}'.format( self._doc_loc, SIDECAR_EXT )
    
    try:
      for name, map in self._sidecarXML.get_maps( ).iteritems( ):
        self._maps[ name ] = ( map )
    except TypeError:
      # No Maps
      new_map = xml_sidecar.Map( )
      new_map.set_name( 'New Map' )
      self._sidecarXML.add_map( new_map )
      self._maps[ map.get_name( ) ] = ( map )
    else:
      for map in self._maps:
        print map
      
    self.grid( )
    self.build( )
    
    self.master.title( 'PSD Export Settings' )
    self.mainloop( )
    
  def build( self ):
    """
    Build the UI for the Export settings UI
    """

    self.top_buttons = Tkinter.Frame( self, padx = 5, pady = 5 )
    self.top_buttons.grid( row = 0, column = 0 )
    
    self.map_widgets = Tkinter.Frame( self, padx = 5, pady = 5 )
    self.map_widgets.grid( row = 1, column = 0 )
    
    self.out_buttons = Tkinter.Frame( self, padx = 5, pady = 5 )
    self.out_buttons.grid( row = 2, column = 0 )
    
    # Top Buttons setup
    self.btn_new_map = Tkinter.Button( self.top_buttons,
                                       text = 'New Map',
                                       command = self._handle_new_map )
    
    self.btn_new_map.grid( row = 0, column = 0 )
    
    self.btn_new_from_temp = Tkinter.Button( self.top_buttons,
                                             text = 'New Map from Template',
                                             command = self._handle_new_from_template )
    
    self.btn_new_from_temp.grid( row = 0, column = 1)
    
    self.btn_set_export_loc = Tkinter.Button( self.top_buttons,
                                              text = 'Set Export Location',
                                              command = self._handle_set_export_loc )
    self.btn_set_export_loc.grid( row = 0, column = 2)
    
    # Map Widgets setup  
    i = 0
    for map in self._maps.values( ):
      widgets.MapWidget( self.map_widgets, self, map ).grid( row = i, column = 0 )
      i = i + 1
      
    # Bottom Buttons setup
    
    self.btn_save_settings = Tkinter.Button( self.out_buttons,
                                             text = 'Save export settings',
                                             command = self._handle_save )
    self.btn_save_settings.grid( row = 0, column = 0 )
    
    self.btn_export = Tkinter.Button( self.out_buttons,
                                       text = 'Save settings and Export',
                                       command = self._handle_save_and_export )
    self.btn_export.grid( row = 0, column = 1 )
  
  def export( self ):
    """
    export all maps.
    """
    
    for map in self._maps.values( ):
      map.export( )
      
  def save_sidecar( self ):
    """
    Save the sidecar file(s)
    """
    self._sidecarXML.set_maps( self._maps )  
    self._sidecarXML.write_settings( )
         
  def _handle_new_map( self ):
    """
    Handle button click for 'new map' button
    """
    pass
  
  def _handle_new_from_template( self ):
    """
    Handle button click for 'New from Template' button.
    """
    pass
  
  def _handle_set_export_loc( self ):
    """
    Handle button click for 'set export location' button.
    """
    
  def _handle_save_and_export( self ):
    """
    Handle button click for 'Save settings and export' button.
    """
    self.save_sidecar( )
    self.export( )
    
  def _handle_save( self ):
    """
    Handle button click for 'Save Export Setttings' button.
    """
    
    self.save_sidecar( )