/*
 * ----------------------------------------------------------------------------- 
 * GDBPlugin - ArcGIS Metadata Plugin
 * ----------------------------------------------------------------------------- 
 *        
 *      Copyright 2011 Provincie Drenthe
 *      
 *      This program is free software; you can redistribute it and/or modify
 *      it under the terms of the GNU General Public License as published by
 *      the Free Software Foundation; either version 2 of the License, or
 *      (at your option) any later version.
 *      
 *      This program is distributed in the hope that it will be useful,
 *      but WITHOUT ANY WARRANTY; without even the implied warranty of
 *      MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
 *      GNU General Public License for more details.
 *      
 *      You should have received a copy of the GNU General Public License
 *      along with this program; if not, write to the Free Software
 *      Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston,
 *      MA 02110-1301, USA.
 */

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using Microsoft.Win32;
using ESRI.ArcGIS.Carto;
using ESRI.ArcGIS.Geoprocessor;
using ESRI.ArcGIS.Location;
using ESRI.ArcGIS.GeocodingTools;
using ESRI.ArcGIS.Geodatabase;
using ESRI.ArcGIS.DataSourcesGDB;
using ESRI.ArcGIS.esriSystem;
using ESRI.ArcGIS.Geometry;
using ESRI.ArcGIS.Controls;

namespace GDBplugin
{
    public partial class Form1 : Form
    {
        private IHookHelper hookHelper;                
        private OleDbConnection mydb;
        private OleDbCommand mycommand;
        private OleDbCommand rsselect;        
        private string connection, pdf, layer; 
        private int totaalmeta;        
        

        public Form1(IHookHelper hookHelper)
        {
            this.hookHelper = hookHelper;
            InitializeComponent();
        }       
        
        void Button1Click(object sender, EventArgs e)
        {
        	if (mydb.State == ConnectionState.Open) {
        		mydb.Close();
        	}
        	this.Close();
        }
        
        void Button2Click(object sender, EventArgs e)
        {
        	HelpForm Help1 = new HelpForm();
        	Help1.ShowDialog();
        }
        
        void Button6Click(object sender, EventArgs e)
        {
        	string waarde, SQL, naam, link;
        	int nummer;
        	
        	if (listBox1.SelectedIndex == -1) {
        		MessageBox.Show("Selecteer eerst een dataset uit de lijst.", "Foutmelding");
        		return;
        	}
        	try {
        		naam = "";
        		nummer = listBox1.SelectedIndex;
        		waarde = listBox1.Items[nummer].ToString();
        		SQL = "SELECT ALT_TITEL FROM DATASET WHERE DATASET_TITEL = '" + waarde + "'";
        		mycommand = new OleDbCommand (SQL, mydb);
        		OleDbDataReader myreader = mycommand.ExecuteReader();
        		try {
        			while (myreader.Read()) {
        				naam = myreader.GetString(0).ToString();        				
        			}
        		}
        		catch (Exception ex){
        			MessageBox.Show (ex.ToString(), "Foutmelding");
        		}
        		finally {
        			myreader.Close();
        		}
        		link = pdf + naam + ".pdf";
        		Process.Start (@link);
        	}
        	catch {
        		MessageBox.Show("Er is geen PDF voor deze dataset aanwezig.", "Foutmelding");        		
        	}
        	
        }
        
        void Form1Load(object sender, EventArgs e)
        {
        	try {
        		RegistryKey regKey = Registry.CurrentUser;
        		regKey = regKey.CreateSubKey("Software\\Provincie Drenthe\\GDBplugin");
        		connection = regKey.GetValue("connection").ToString() ;        		
        		pdf = regKey.GetValue("pdf").ToString();
        		layer = regKey.GetValue("layer").ToString();
        	}
        	catch {
        		MessageBox.Show("Er zijn geen GDB Plugin gegevens in het register gevonden.", "Foutmelding");
        		this.Close();
        	}
        	        	
        	//if (database != "" && user != "" && password != "") {
        		if (connection != "") {
        		mydb = new OleDbConnection();        		
        		mydb.ConnectionString = connection;
        		try {
        			mydb.Open();
        			vullenmetadatalijst();
        		}
        		catch {
        			MessageBox.Show("Er is geen verbinding met de database.", "Foutmelding");        			
        			this.Close();
        		}
        	}
        	else {
        		MessageBox.Show("Er zijn geen GDB Plugin gegevens in het register gevonden.", "Foutmelding");
        		this.Close();
        	}
        }
        
        private void vullenmetadatalijst() {
        	String SQL, lijst;
        	
        	SQL = "SELECT DATASET_TITEL FROM DATASET WHERE TYPE = 1 ORDER BY DATASET_TITEL";
        	listBox1.Items.Clear();
        	mycommand = new OleDbCommand (SQL, mydb);
        	OleDbDataReader myreader = mycommand.ExecuteReader();
        	totaalmeta = 0;
        	try {
        		while (myreader.Read()) {
        			totaalmeta = totaalmeta + 1;
        			lijst = myreader.GetString(0).ToString();
        			listBox1.Items.Add (lijst);        			
        		}
        	}        	
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {        		
        		myreader.Close ();
        	}
        	label1.Text = "Dataset (" + totaalmeta.ToString() + " van " + totaalmeta.ToString() + ")";
        }
        
        void vullenlistbox (string SQL, int waarde, ListBox list) {
        	string lijst;
        	
        	list.Items.Clear();
        	rsselect = new OleDbCommand (SQL, mydb);
        	OleDbDataReader myrs = rsselect.ExecuteReader();
        	
        	try {
        		while (myrs.Read()) {
        			lijst = myrs.GetString(waarde).ToString();
        			list.Items.Add (lijst);       			
        		}
        	}        	
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {        		
        		myrs.Close ();
        	}
        }
        	
              
        void Form1FormClosed(object sender, FormClosedEventArgs e)
        {
        	if (mydb.State == ConnectionState.Open) {
        		mydb.Close();
        	}
        }
        
        void Button3Click(object sender, EventArgs e)
        {
        	listBox1.Items.Clear();
        	vullenmetadatalijst();
        }
        
        void Button5Click(object sender, EventArgs e)
        {
        	string waarde, SQL, naam, status, map;
        	int nummer;
        	IEnvelope env;
        	IActiveView pActiveView;
        	IPoint punt;
        	IMap pMap;
        	PropertySet pPropertySet;
        	IWorkspaceFactory pSdeFact;
        	IWorkspace pWorkspace;
        	IFeatureWorkspace pFeatureWorkspace;
        	IFeatureLayer pFeatureLayer;
        	IFeatureClass pFeatureClass;        	
        	ESRI.ArcGIS.Catalog.IGxLayer pGxLayer;
        	ESRI.ArcGIS.Catalog.IGxFile pGxFile;
        	double minx, miny, maxx, maxy;
        	
        	if (listBox1.SelectedIndex == -1) {
        		MessageBox.Show("Selecteer eerst een dataset uit de lijst.", "Foutmelding");
        		return;
        	}
        	try {
        		this.Cursor = Cursors.WaitCursor;
        		naam = "";
        		status = "";
        		nummer = listBox1.SelectedIndex;
        		waarde = listBox1.Items[nummer].ToString();
        		SQL = "SELECT NAAM, FYSIEKE_LOCATIE FROM DATASET WHERE DATASET_TITEL = '" + waarde + "'";
        		mycommand = new OleDbCommand (SQL, mydb);
        		OleDbDataReader myreader = mycommand.ExecuteReader();
        		try {
        			while (myreader.Read()) {
        				naam = myreader.GetString(0).ToString();
        				status = myreader.GetString(1).ToString();
        			}
        		}
        		catch (Exception ex){
        			MessageBox.Show (ex.ToString(), "Foutmelding");
        			this.Cursor = Cursors.Default;
        		}
        		finally {
        			myreader.Close();
        		}
        		
        		map = layer + status + "\\" + naam + ".lyr";        		
        		
        		try{
        			if (System.IO.File.Exists(@map)) {        				
        				pActiveView = hookHelper.ActiveView;
        				env = pActiveView.Extent;
        				punt = env.LowerLeft;
        				minx = punt.X;
        				miny = punt.Y;
        				
        				punt = env.UpperRight;
        				maxx = punt.X;
        				maxy = punt.Y;
        				
        				pGxLayer = new ESRI.ArcGIS.Catalog.GxLayerClass();
        				pGxFile = (ESRI.ArcGIS.Catalog.IGxFile)pGxLayer;
        				pGxFile.Path = map;
        				
        				pMap = hookHelper.FocusMap;
        				pMap.AddLayer(pGxLayer.Layer);
        				
        				pActiveView.Extent.XMin = minx;
        				pActiveView.Extent.XMax = maxx;
        				pActiveView.Extent.YMin = miny;
        				pActiveView.Extent.YMax = maxy;
        				
        				pActiveView.Refresh();
        				this.Cursor = Cursors.Default;
        			}
        			else {
        				MessageBox.Show("Er is geen opmaakbestand van de geselecteerde dataset aanwezig. De dataset wordt nu zonder opmaak uit de geo-database gehaald.","Informatie");
        				SQL = "SELECT ALT_TITEL FROM DATASET WHERE DATASET_TITEL = '" + waarde + "'";
        				mycommand = new OleDbCommand (SQL, mydb);
        				OleDbDataReader myreader1 = mycommand.ExecuteReader();
        				try {
        					while (myreader1.Read()) {
        						pActiveView = hookHelper.ActiveView;
        						
        						pPropertySet = new PropertySetClass();
        						pPropertySet.SetProperty("SERVER", "chios");
        						pPropertySet.SetProperty("INSTANCE", "5151");
        						pPropertySet.SetProperty("DATABASE", "");
        						pPropertySet.SetProperty("USER", "gisuser");
        						pPropertySet.SetProperty("PASSWORD", "zonnetje");
        						pPropertySet.SetProperty("VERSION", "SDE.DEFAULT");
        						
        						env = pActiveView.Extent;
        						punt = env.LowerLeft;        						
		        				minx = punt.X;
		        				miny = punt.Y;
		        				
		        				punt = env.UpperRight;
		        				maxx = punt.X;
		        				maxy = punt.Y;
		        				
		        				pSdeFact = new SdeWorkspaceFactoryClass();
		        				pWorkspace = pSdeFact.Open(pPropertySet, 0);		        				
		        				pFeatureWorkspace = pWorkspace as IFeatureWorkspace; 
		        				
		        				pFeatureClass = pFeatureWorkspace.OpenFeatureClass(myreader1.GetString(0).ToString());
		        				
		        				pMap = hookHelper.FocusMap;
		        				
		        				pFeatureLayer = new FeatureLayerClass();
		        				pFeatureLayer.FeatureClass = pFeatureClass;
		        				
		        				pFeatureLayer.Name = pFeatureClass.AliasName;
		        				pFeatureLayer.Visible = true;
		        				pFeatureLayer.Selectable = true;
		        				pMap.AddLayer(pFeatureLayer);
		        				
		        				pActiveView.Extent.XMin = minx;
		        				pActiveView.Extent.XMax = maxx;
		        				pActiveView.Extent.YMin = miny;
		        				pActiveView.Extent.YMax = maxy;
		        				
		        				pActiveView.Refresh();	
		        				this.Cursor = Cursors.Default;		        		
        					}
        				}
        				catch (Exception ex){
        					MessageBox.Show (ex.ToString(), "Foutmelding");
        					this.Cursor = Cursors.Default;
        				}
        				finally {
        					myreader.Close();        					
        				} 	
        				
        			}
        			
        		}
        		catch (Exception ex) {
        			MessageBox.Show (ex.ToString(), "Foutmelding");
        			this.Cursor = Cursors.Default;
        		}
        	}
        	catch {
        		MessageBox.Show("Er is geen PDF voor deze dataset aanwezig.", "Foutmelding");  
        		this.Cursor = Cursors.Default;
        	}	
        }
        
        void ListBox1DoubleClick(object sender, EventArgs e)
        {
        	int nummer, x;
        	string waarde, SQL, datacode, stditem, metapersoon, geoloket;
        	OleDbCommand mycommand1;
        	
        	this.Cursor = Cursors.WaitCursor;
        	//naam = "";
        	nummer = listBox1.SelectedIndex;
        	waarde = listBox1.Items[nummer].ToString(); 
        	datacode = "";
        	
        	
        	SQL = "SELECT a.DATASET_TITEL, a.ALT_TITEL, b.TEKST, a.OPMERKING, a.OPBOUWDATUM, a.BRONDATUM, c.BRONVERMELDING, " +
			"c.OPBOUWMETHODE , a.ACTIE, a.STATUS FROM DATASET a, MEMOTABEL b, GEOGRAFISCH c WHERE c.DATACODE = a.DATACODE AND " + 
			"a.OMSCHRIJVING_CODE = b.Code AND a.DATASET_TITEL = '" + waarde + "'";
        	       	
        	mycommand = new OleDbCommand (SQL, mydb);
        	OleDbDataReader myreader = mycommand.ExecuteReader();
        	try {
        		while (myreader.Read()) {
        			if (myreader.GetValue(0) == DBNull.Value ) {
        				textBox2.Text = "";
        			}
        			else {
        				textBox2.Text = myreader.GetString(0).ToString();        				
        			}
        			if (myreader.GetValue(1) == DBNull.Value ) {
        				textBox3.Text = "";
        			}
        			else {
        				textBox3.Text = myreader.GetString(1).ToString();        				
        			}
        			if (myreader.GetValue(2) == DBNull.Value ) {
        				textBox4.Text = "";
        			}
        			else {
        				textBox4.Text = myreader.GetString(2).ToString();        				
        			}
        			if (myreader.GetValue(3) == DBNull.Value ) {
        				textBox5.Text = "";
        			}
        			else {        				
        				textBox5.Text = myreader.GetString(3).ToString();
        			}
        			if (myreader.GetValue(4) == DBNull.Value ) {
        				textBox6.Text = "";
        			}
        			else {        				
        				textBox6.Text = myreader.GetDateTime(4).ToShortDateString();        				
        			}
        			if (myreader.GetValue(5) == DBNull.Value ) {
        				textBox7.Text = "";
        			}
        			else {
        				textBox7.Text = myreader.GetDateTime(5).ToShortDateString();        				
        			}
        			if (myreader.GetValue(6) == DBNull.Value ) {
        				textBox8.Text = "";
        			}
        			else {
        				textBox8.Text = myreader.GetString(6).ToString();        				
        			}
        			if (myreader.GetValue(7) == DBNull.Value ) {
        				textBox9.Text = "";
        			}
        			else {
        				textBox9.Text = myreader.GetString(7).ToString();        				
        			}
        			if (myreader.GetValue(8) == DBNull.Value ) {
        				textBox10.Text = "";
        			}
        			else {
        				textBox10.Text = myreader.GetString(8).ToString();        				
        			}
        			if (myreader.GetValue(9) == DBNull.Value ) {
        				textBox11.Text = "";
        			}
        			else {
        				textBox11.Text = myreader.GetString(9).ToString();        				
        			}        			
        		}        		
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {        		
        		myreader.Close();
        	}
        	
        	
        	SQL = "SELECT a.DATACODE, a.BELEIDSVELD, a.VEILIGHEID, a.TEAM, a.THEMA, a.GEBRUIKSBEPERKING, a.JURIDISCH, a.COPYRIGHT, a.BIJHOUDING, b.SCHAAL, " + 
			"a.CONTACT_LEVERANCIER, a.DEKKING_BEGIN, a.DEKKING_EIND, c.CONTACTPERSOON FROM DATASET a, GEOGRAFISCH b, CONTACT c WHERE a.DATACODE " + 
			"= b.DATACODE AND a.CONTACTPERSOON = c.CONTACT_ID AND a.DATASET_TITEL = '" + waarde + "'";       	
        	       	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	try {
        		while (myreader.Read()) {        			
        			datacode = myreader.GetValue(0).ToString();
        			
        			if (myreader.GetValue(13) == DBNull.Value ) {
        				textBox12.Text = "";
        			}
        			else {
        				textBox12.Text = myreader.GetString(13).ToString();        				
        			}
        			
        			if (myreader.GetValue(1) == DBNull.Value ) {
        				textBox13.Text = "";
        			}
        			else {
        				textBox13.Text = myreader.GetString(1).ToString();        				
        			}
        			
        			if (myreader.GetValue(3) == DBNull.Value ) {
        				textBox14.Text = "";
        			}
        			else {
        				textBox14.Text = myreader.GetString(3).ToString();        				
        			}
        			
        			if (myreader.GetValue(4) == DBNull.Value ) {
        				textBox15.Text = "";
        			}
        			else {        				
        				textBox15.Text = myreader.GetString(4).ToString();
        			}
        			
        			if (myreader.GetValue(5) == DBNull.Value ) {
        				textBox16.Text = "";
        			}
        			else {        				
        				textBox16.Text = myreader.GetString(5).ToString();        				
        			}
        			
        			if (myreader.GetValue(2) == DBNull.Value ) {
        				textBox17.Text = "";
        			}
        			else {
        				textBox17.Text = myreader.GetString(2).ToString();        				
        			}
        			
        			if (myreader.GetValue(6) == DBNull.Value ) {
        				textBox18.Text = "";
        			}
        			else {
        				textBox18.Text = myreader.GetString(6).ToString();        				
        			}
        			
        			if (myreader.GetValue(7) == DBNull.Value ) {
        				textBox19.Text = "";
        			}
        			else {
        				textBox19.Text = myreader.GetString(7).ToString();        				
        			}        			
        			
        			if (myreader.GetValue(8) == DBNull.Value ) {
        				textBox20.Text = "";
        			}
        			else {
        				textBox20.Text = myreader.GetString(8).ToString();        				
        			}
        			
        			if (myreader.GetValue(9) == DBNull.Value ) {
        				textBox21.Text = "";
        			}
        			else {
        				textBox21.Text = myreader.GetString(9).ToString();        				
        			}
        			
        			if (myreader.GetValue(10) == DBNull.Value ) {
        				textBox22.Text = "";
        			}
        			else {
        				textBox22.Text = myreader.GetString(10).ToString();        				
        			}
        			
        			if (myreader.GetValue(11) == DBNull.Value ) {
        				textBox24.Text = "";
        			}
        			else {
        				textBox24.Text = myreader.GetDateTime(11).ToShortDateString();   				
        			}
        			
        			if (myreader.GetValue(12) == DBNull.Value ) {
        				textBox25.Text = "";
        			}
        			else {
        				textBox25.Text = myreader.GetDateTime(12).ToShortDateString();         				
        			}        			
        		}        		
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}
        	
        	
        	SQL = "SELECT a.TREFWOORD FROM TREFTEXT a, TREFCODE b WHERE " +
   			"b.TREFCODE = a.TREFCODE AND b.DATACODE = " + datacode + " ORDER BY a.TREFWOORD";       	
        	       	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	x = 0;
        	try {
        		while (myreader.Read()) {
        			if (x == 0) {
        				textBox23.Text = myreader.GetString(0).ToString();
        				x = 1;
        			}
        			else {
        				textBox23.Text = textBox23.Text + ", " + myreader.GetString(0).ToString();
        			}
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}   
        	
        	SQL = "SELECT b.DEELGEBIED, a.RSCHEMA, a.AANVUL_INFO, a.NAAM, a.FYSIEKE_LOCATIE, a.DATATYPE, " + 
			"b.POS_NAUWKEURIGHEID , a.KWALITEIT_BESCH, b.GEOMETRIE " + 
			"FROM DATASET a, GEOGRAFISCH b WHERE a.DATACODE = b.DATACODE AND a.DATASET_TITEL = '" + waarde + "'";
        	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	try {
        		while (myreader.Read()) { 
        			if (myreader.GetValue(0) == DBNull.Value ) {
        				textBox26.Text = "";
        			}
        			else {
        				textBox26.Text = myreader.GetString(0).ToString();        				
        			}
        			
        			if (myreader.GetValue(1) == DBNull.Value ) {
        				textBox27.Text = "";
        			}
        			else {
        				textBox27.Text = myreader.GetString(1).ToString();        				
        			}
        			
        			if (myreader.GetValue(2) == DBNull.Value ) {
        				textBox28.Text = "";
        			}
        			else {
        				textBox28.Text = myreader.GetString(2).ToString();        				
        			}
        			
        			if (myreader.GetValue(3) == DBNull.Value ) {
        				textBox29.Text = "";
        			}
        			else {        				
        				textBox29.Text = myreader.GetString(3).ToString();
        			}
        			
        			if (myreader.GetValue(4) == DBNull.Value ) {
        				textBox30.Text = "";
        			}
        			else {        				
        				textBox30.Text = myreader.GetString(4).ToString();        				
        			}
        			
        			if (myreader.GetValue(5) == DBNull.Value ) {
        				textBox31.Text = "";
        			}
        			else {
        				textBox31.Text = myreader.GetString(5).ToString();        				
        			}
        			
        			if (myreader.GetValue(8) == DBNull.Value ) {
        				textBox32.Text = "";
        			}
        			else {
        				textBox32.Text = myreader.GetString(8).ToString();        				
        			}
        			
        			if (myreader.GetValue(6) == DBNull.Value ) {
        				textBox33.Text = "";
        			}
        			else {
        				textBox33.Text = myreader.GetString(6).ToString();        				
        			}        			
        			
        			if (myreader.GetValue(7) == DBNull.Value ) {
        				textBox34.Text = "";
        			}
        			else {
        				textBox34.Text = myreader.GetString(7).ToString();        				
        			}
        			
        			if (myreader.GetValue(0) == DBNull.Value ) {
        				textBox35.Text = "";
        				textBox36.Text = "";
        				textBox37.Text = "";
        				textBox38.Text = "";
			        				
        			}
        			else {
        				SQL = "SELECT MIN_X, MAX_X, MIN_Y, MAX_Y FROM GEBIED WHERE GEBIED = '" + myreader.GetString(0).ToString() + "'";
        	       	
			        	mycommand1 = new OleDbCommand (SQL, mydb);
			        	OleDbDataReader myreader1 = mycommand1.ExecuteReader();
			        	x = 0;
			        	try {
			        		while (myreader1.Read()) {
			        			textBox35.Text = myreader1.GetString(0).ToString();
        						textBox36.Text = myreader1.GetString(1).ToString();
        						textBox37.Text = myreader1.GetString(2).ToString();
        						textBox38.Text = myreader1.GetString(3).ToString();
			        		}
			        	}
			        	catch (Exception ex){
			        		MessageBox.Show (ex.ToString(), "Foutmelding");
			        	}
			        	finally {
			        		myreader1.Close();
			        	}   
        			}
        		}        		
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}
        	
        	SQL = "SELECT a.DATACODE, b.STD_ITEM FROM DATASET a, GEOGRAFISCH b WHERE " +
			"a.DATACODE = b.DATACODE AND a.DATASET_TITEL = '" + waarde + "'";
        	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	stditem = "";
        	
        	try {
        		while (myreader.Read()) { 
        			if (myreader.GetValue(0) == DBNull.Value ) {        				
        			}
        			else {
        				datacode = myreader.GetValue(0).ToString();        				
        			}
        			
        			if (myreader.GetValue(1) == DBNull.Value ) {
        			}
        			else {        				
        				stditem = myreader.GetString(1).ToString();
        			}
        			
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}
        	
        	if (stditem == "") {
        		textBox39.Text = "";
        	}
        	else {
        		textBox39.Text = stditem;
        	}
        	
        	if (datacode != "") {
        		vullenlistbox("SELECT ITEMNAAM FROM ITEMS WHERE DATACODE = " + datacode + " ORDER BY VOLGNR", 0, listBox2);
        	}
        	
        	SQL = "SELECT a.METAPERSOON, a.OPBOUWDATUM, a.METADATASTD, a.TAAL, a.KARAKTERSET, a.VERSIE_METASTD, a.CODE_REF, a.ORG_NAMESPACE, " + 
			"a.GEOLOKET FROM DATASET a WHERE a.DATASET_TITEL = '" + waarde + "'";
        	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();        	
        	
        	metapersoon = "";
        	geoloket = "";
        	
        	try {
        		while (myreader.Read()) { 
        			if (myreader.GetValue(0) == DBNull.Value ) {         				
        			}
        			else {
        				metapersoon = myreader.GetValue(0).ToString();
        			}
        			
        			if (myreader.GetValue(1) == DBNull.Value ) {
        				textBox45.Text = "";
        			}
        			else {
        				textBox45.Text = myreader.GetDateTime(1).ToShortDateString();
        			}
        			
        			if (myreader.GetValue(3) == DBNull.Value ) {
        				textBox46.Text = "";
        			}
        			else {
        				textBox46.Text = myreader.GetString(3).ToString();
        			}
        			
        			if (myreader.GetValue(4) == DBNull.Value ) {
        				textBox47.Text = "";
        			}
        			else {
        				textBox47.Text = myreader.GetString(4).ToString();
        			}
        			
        			if (myreader.GetValue(2) == DBNull.Value ) {
        				textBox48.Text = "";
        			}
        			else {
        				textBox48.Text = myreader.GetString(2).ToString();
        			}
        			
        			if (myreader.GetValue(5) == DBNull.Value ) {
        				textBox49.Text = "";
        			}
        			else {
        				textBox49.Text = myreader.GetString(5).ToString();
        			}
        			
        			if (myreader.GetValue(6) == DBNull.Value ) {
        				textBox50.Text = "";
        			}
        			else {
        				textBox50.Text = myreader.GetString(6).ToString();
        			}
        			
        			if (myreader.GetValue(7) == DBNull.Value ) {
        				textBox51.Text = "";
        			}
        			else {
        				textBox51.Text = myreader.GetString(7).ToString();
        			}
        			
        			if (myreader.GetValue(8) == DBNull.Value ) {        				
        			}
        			else {
        				geoloket = myreader.GetValue(8).ToString();
        			}        			
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}  	
        	
        	
        	SQL = "SELECT CONTACTPERSOON FROM CONTACT WHERE CONTACT_ID = '" + metapersoon + "'";
        	       	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	
        	try {
        		while (myreader.Read()) {
        			textBox40.Text = myreader.GetString(0).ToString();
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}   
        	
        	SQL = "SELECT CONTACTPERSOON FROM CONTACT WHERE CONTACT_ID = '" + geoloket + "'";
        	       	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	
        	try {
        		while (myreader.Read()) {
        			textBox52.Text = myreader.GetString(0).ToString();
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}
        	this.Cursor = Cursors.Default;
        	
        }
        
        private void opschonen() {
        	textBox2.Text = "";
        	textBox3.Text = "";
        	textBox4.Text = "";
        	textBox5.Text = "";
        	textBox6.Text = "";
        	textBox7.Text = "";
        	textBox8.Text = "";
        	textBox9.Text = "";
        	textBox10.Text = "";
        	textBox11.Text = "";
        	textBox12.Text = "";
        	textBox13.Text = "";
        	textBox14.Text = "";
        	textBox15.Text = "";
        	textBox16.Text = "";
        	textBox17.Text = "";
        	textBox18.Text = "";
        	textBox19.Text = "";
        	textBox20.Text = "";
        	textBox21.Text = "";
        	textBox22.Text = "";
        	textBox23.Text = "";
        	textBox24.Text = "";
        	textBox25.Text = "";       	
        	textBox26.Text = "";
        	textBox27.Text = "";
        	textBox28.Text = "";
        	textBox29.Text = "";
        	textBox30.Text = "";
        	textBox31.Text = "";
        	textBox32.Text = "";
        	textBox33.Text = "";
        	textBox34.Text = "";
        	textBox35.Text = "";
        	textBox36.Text = "";
        	textBox37.Text = "";
        	textBox38.Text = ""; 
        	textBox39.Text = "";
        	listBox2.Items.Clear();
        	textBox41.Text = "";
        	textBox42.Text = "";
        	textBox43.Text = "";
			textBox44.Text = "";
			textBox45.Text = "";
        	textBox46.Text = "";
        	textBox47.Text = "";
			textBox48.Text = "";
			textBox49.Text = "";
        	textBox50.Text = "";
        	textBox51.Text = "";
			textBox52.Text = "";
        }  
        
        void ListBox2DoubleClick(object sender, EventArgs e)
        {
        	String SQL, waarde, waarde1, dc;        	
        	int nummer;
        	nummer = listBox1.SelectedIndex;
        	waarde1 = listBox1.Items[nummer].ToString();
        	dc = "";
        	
        	SQL = "SELECT DATACODE FROM DATASET WHERE DATASET_TITEL = '" + waarde1 + "'";
        	       	
        	mycommand = new OleDbCommand (SQL, mydb);
        	OleDbDataReader myreader = mycommand.ExecuteReader();
        	
        	try {
        		while (myreader.Read()) {
        			dc = myreader.GetValue(0).ToString();
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}
        	
        	nummer = listBox2.SelectedIndex;
        	waarde = listBox2.Items[nummer].ToString();
        	
        	SQL = "SELECT a.VOLGNR, a.ITEMNAAM, a.ITEMDEFINITIE, a.EENHEID, b.TEKST FROM ITEMS a, MEMOTABEL b WHERE a.DOMEIN = b.CODE " +
        		"AND a.DATACODE = '" + dc + "' AND a.ITEMNAAM = '" + waarde + "'";
        	
        	mycommand = new OleDbCommand (SQL, mydb);
        	myreader = mycommand.ExecuteReader();
        	
        	try {
        		while (myreader.Read()) {
        			if (myreader.GetValue(1) == DBNull.Value ) {
        				textBox41.Text = "";
        			}
        			else {
        				textBox41.Text = myreader.GetString(1).ToString();
        			}
        			
        			if (myreader.GetValue(2) == DBNull.Value ) {
        				textBox42.Text = "";
        			}
        			else {
        				textBox42.Text = myreader.GetString(2).ToString();
        			}
        			
        			if (myreader.GetValue(3) == DBNull.Value ) {
        				textBox43.Text = "";
        			}
        			else {
        				textBox43.Text = myreader.GetString(3).ToString();
        			}
        			
        			if (myreader.GetValue(4) == DBNull.Value ) {
        				textBox44.Text = "";
        			}
        			else {
        				textBox44.Text = myreader.GetString(4).ToString();
        			}
        		}
        	}
        	catch (Exception ex){
        		MessageBox.Show (ex.ToString(), "Foutmelding");
        	}
        	finally {
        		myreader.Close();
        	}
        }
        
        void TextBox1KeyPress(object sender, KeyPressEventArgs e)
        {
        	int totaalset;
        	string SQL, lijst, waarde, waarde1;
        	
        	if (e.KeyChar == (char)Keys.Enter ) {
        		if (textBox1.Text != "") {
        			opschonen();
        			waarde = textBox1.Text;
        			waarde1 = waarde.ToUpper();
        			SQL = "SELECT DISTINCT DATASET.DATACODE, DATASET.DATASET_TITEL FROM ((DATASET INNER JOIN MEMOTABEL " + 
        			"ON DATASET.OMSCHRIJVING_CODE = MEMOTABEL.CODE) INNER JOIN TREFCODE ON DATASET.DATACODE = TREFCODE.DATACODE) " + 
        			"INNER JOIN TREFTEXT ON TREFCODE.TREFCODE = TREFTEXT.TREFCODE WHERE " +
        				"DATASET.DATASET_TITEL LIKE '%" + waarde  + "%' OR MEMOTABEL.TEKST Like '%" + waarde + "%' OR TREFTEXT.TREFWOORD " +
        			"LIKE '%" + waarde + "%' OR DATASET.DATASET_TITEL LIKE '%" + waarde1 +  "%' OR MEMOTABEL.TEKST Like '%" + waarde1 + 
        			"%' OR TREFTEXT.TREFWOORD LIKE '%" + waarde1 + "%' AND DATASET.TYPE = 1 ORDER BY DATASET.DATASET_TITEL";
        			
        			listBox1.Items.Clear();
        			
        			mycommand = new OleDbCommand (SQL, mydb);
		        	OleDbDataReader myreader = mycommand.ExecuteReader();
		        	totaalset = 0;
		        	try {
		        		while (myreader.Read()) {
		        			totaalset = totaalset + 1;
		        			lijst = myreader.GetString(1).ToString();
		        			listBox1.Items.Add (lijst);        			
		        		}
		        	}        	
		        	catch (Exception ex){
		        		MessageBox.Show (ex.ToString(), "Foutmelding");
		        	}
		        	finally {        		
		        		myreader.Close ();
		        	}
		        	label1.Text = "Dataset (" + totaalset.ToString() + " van " + totaalmeta.ToString() + ")";
        		}
        	}
        }
        
        void Button4Click(object sender, EventArgs e)
        {
        	int totaalset;
        	string SQL, lijst, waarde, waarde1;        	
        	
        		if (textBox1.Text != "") {
        			opschonen();
        			waarde = textBox1.Text; 
        			waarde1 = waarde.ToUpper();
        			SQL = "SELECT DISTINCT DATASET.DATACODE, DATASET.DATASET_TITEL FROM ((DATASET INNER JOIN MEMOTABEL " + 
        			"ON DATASET.OMSCHRIJVING_CODE = MEMOTABEL.CODE) INNER JOIN TREFCODE ON DATASET.DATACODE = TREFCODE.DATACODE) " + 
        			"INNER JOIN TREFTEXT ON TREFCODE.TREFCODE = TREFTEXT.TREFCODE WHERE " +
        				"DATASET.DATASET_TITEL LIKE '%" + waarde  + "%' OR MEMOTABEL.TEKST Like '%" + waarde + "%' OR TREFTEXT.TREFWOORD " +
        			"LIKE '%" + waarde + "%' OR DATASET.DATASET_TITEL LIKE '%" + waarde1 +  "%' OR MEMOTABEL.TEKST Like '%" + waarde1 + 
        			"%' OR TREFTEXT.TREFWOORD LIKE '%" + waarde1 + "%' AND DATASET.TYPE = 1 ORDER BY DATASET.DATASET_TITEL";      			
        			
        			listBox1.Items.Clear();
        			
        			mycommand = new OleDbCommand (SQL, mydb);
		        	OleDbDataReader myreader = mycommand.ExecuteReader();
		        	totaalset = 0;
		        	try {
		        		while (myreader.Read()) {
		        			totaalset = totaalset + 1;
		        			lijst = myreader.GetString(1).ToString();
		        			listBox1.Items.Add (lijst);        			
		        		}
		        	}        	
		        	catch (Exception ex){
		        		MessageBox.Show (ex.ToString(), "Foutmelding");
		        	}
		        	finally {        		
		        		myreader.Close ();
		        	}
		        	label1.Text = "Dataset (" + totaalset.ToString() + " van " + totaalmeta.ToString() + ")";
        		}
        		else {
        			MessageBox.Show ("Type eerst een zoekterm in.", "Foutmelding");
        		}        	
        }
    }
}