﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace FIA_Biosum_Manager
{
  public class frmFCSTreeVolumeEdit : Form
  {
    const int COL_VOLCFGRS = 0;
    const int COL_VOLCSGRS = 1;
    const int COL_VOLCFNET = 2;
    const int COL_DRYBIOM = 3;
    const int COL_DRYBIOT = 4;
    const int COL_ID = 5;
    const int COL_STATE = 6;
    const int COL_COUNTY = 7;
    const int COL_PLOT = 8;
    const int COL_FVSVARIANT = 9;
    const int COL_INVYR = 10;
    const int COL_TREEID = 11;
    const int COL_VOLLOCGRP = 14;
    const int COL_SPCD = 11;
    const int COL_DBH = 12;
    const int COL_HT = 13;
    const int COL_ACTUALHT = 15;
    const int COL_CR = 18;
    const int COL_STATUSCD = 16;
    const int COL_TREECLCD = 17;
    const int COL_CULL = 19;
    const int COL_ROUGHCULL = 20;
    const int COL_TRECN = 22;
    const int COL_PLTCN = 23;
    const int COL_CNDCN = 24;
    private System.Windows.Forms.Button btnTreeVolBatch;
    private System.Windows.Forms.Panel panel1;
    private System.Windows.Forms.GroupBox groupBox1;
    private System.Windows.Forms.GroupBox groupBox2;
    private System.Windows.Forms.Label label7;
    private System.Windows.Forms.TextBox txtDbh;
    private System.Windows.Forms.Label label6;
    private System.Windows.Forms.TextBox txtSpCd;
    private System.Windows.Forms.Label label5;
    private System.Windows.Forms.TextBox txtInvYr;
    private System.Windows.Forms.Label label4;
    private System.Windows.Forms.TextBox txtVolLocGrp;
    private System.Windows.Forms.Label label3;
    private System.Windows.Forms.TextBox txtStateCd;
    private System.Windows.Forms.Label label2;
    private System.Windows.Forms.TextBox txtPlot;
    private System.Windows.Forms.Label label1;
    private System.Windows.Forms.TextBox txtCountyCd;
    private System.Windows.Forms.Button btnTreeVolSingle;
    private System.Windows.Forms.Label label9;
    private System.Windows.Forms.TextBox txtActualHt;
    private System.Windows.Forms.Label label8;
    private System.Windows.Forms.TextBox txtHt;
    private System.Windows.Forms.Label label10;
    private System.Windows.Forms.TextBox txtCR;
    private System.Windows.Forms.Label label12;
    private System.Windows.Forms.TextBox txtTreeClCd;
    private System.Windows.Forms.Label label11;
    private System.Windows.Forms.TextBox txtStatusCd;
    private System.Windows.Forms.GroupBox groupBox3;
    private System.Windows.Forms.Label lblDRYBIOM;
    private System.Windows.Forms.Label label20;
    private System.Windows.Forms.Label lblDRYBIOT;
    private System.Windows.Forms.Label label19;
    private System.Windows.Forms.Label lblVOLCFNET;
    private System.Windows.Forms.Label label18;
    private System.Windows.Forms.Label lblVOLCSGRS;
    private System.Windows.Forms.Label label17;
    private System.Windows.Forms.Label lblVOLCFGRS;
    private System.Windows.Forms.Label label15;
    private System.Windows.Forms.Label label14;
    private System.Windows.Forms.TextBox txtRoughCull;
    private System.Windows.Forms.Label label13;
    private System.Windows.Forms.TextBox txtCull;
    private System.Windows.Forms.Button btnDefaultSingle;
    private System.Windows.Forms.Button btnLinkTableTest;
    //
    //GRID POP UP MENU
    //
    private const int MENU_FILTERBYVALUE = 0;
    private const int MENU_FILTERBYENTEREDVALUE = 1;
    private const int MENU_REMOVEFILTER = 2;
    private const int MENU_UNIQUEVALUES = 3;
    private const int MENU_MODIFY = 5;
    private const int MENU_DELETE = 6;
    private const int MENU_SELECTALL = 8;
    private const int MENU_IDXASC = 10;
    private const int MENU_IDXDESC = 11;
    private const int MENU_REMOVEIDX = 12;
    private const int MENU_MAX = 14;
    private const int MENU_MIN = 15;
    private const int MENU_AVG = 16;
    private const int MENU_SUM = 17;
    private const int MENU_COUNTBYVALUE = 18;
    private Button btnEdit;
    private Button btnLoad;
    private ComboBox cmbDatasource;
    private uc_gridview uc_gridview1 = new uc_gridview();
    private Button btnCancel;
    private Label lblPerc = new Label();
    private RxTools m_oRxTools = new RxTools();
    private Queries m_oQueries = new Queries();
    string[] m_strFVSTreeTableLinkNameArray = null;
    string m_strTempDBFile = "";
    string m_strGridTableSource = "";
    static string m_strOldPerc = "0";
    private ProgressBarBasic.ProgressBarBasic progressBarBasic1;

    

    private ado_data_access m_oAdo = new ado_data_access();
    public frmFCSTreeVolumeEdit()
    {
     this.SetStyle(ControlStyles.OptimizedDoubleBuffer |
               ControlStyles.AllPaintingInWmPaint,
               true);
      int x,y;
      string strFVSTreeTableLinkNameList="";
      string strTableName="";

      dao_data_access oDao = new dao_data_access();
      InitializeComponent();
      groupBox1.Controls.Add(uc_gridview1);
      uc_gridview1.Top = 15; uc_gridview1.Left = 5;
      uc_gridview1.Width = groupBox1.Width - 10;
      uc_gridview1.Height = cmbDatasource.Top - uc_gridview1.Top - 5;
     
           
      uc_gridview1.CloseButton_Visible = false;
      uc_gridview1.Show();
      ResizeForm();
      
      if (frmMain.g_oFrmMain.frmProject.uc_project1.txtRootDirectory.Text.Trim().Length > 0)
      {

          m_oQueries.m_oFvs.LoadDatasource = true;
          m_oQueries.m_oFIAPlot.LoadDatasource = true;
          m_oQueries.LoadDatasources(true);
          m_strTempDBFile = m_oQueries.m_strTempDbFile;
          oDao.CreateTableLink(m_strTempDBFile, "treesample", frmMain.g_oEnv.strAppDir + "\\db\\treesample.mdb", "treesample");
          oDao.m_DaoWorkspace.Close();
          //
          //CREATE LINK IN TEMP MDB TO ALL VARIANT CUTLIST TABLES
          //
          m_oRxTools.CreateTableLinksToFVSOutTreeListTables(m_oQueries, m_oQueries.m_strTempDbFile);
          //
          //OPEN CONNECTION TO TEMP DB FILE
          //
          m_oAdo = new ado_data_access();
          m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_oQueries.m_strTempDbFile, "", ""));

          if (m_oAdo.m_intError == 0)
          {
              RxPackageItem_Collection oRxPackageItemCollection = new RxPackageItem_Collection();
              m_oRxTools.LoadAllRxPackageItemsFromTableIntoRxPackageCollection(m_oAdo, m_oAdo.m_OleDbConnection, m_oQueries, oRxPackageItemCollection);
              //
              //GET LIST OF VARIANTS
              //
              string strVariantsList = m_oRxTools.GetListOfFVSVariantsInPlotTable(m_oAdo, m_oAdo.m_OleDbConnection, m_oQueries.m_oFIAPlot.m_strPlotTable);
              string[] strVariantsArray = frmMain.g_oUtils.ConvertListToArray(strVariantsList, ",");
              //find the variants that have tree cut list tables
              for (x = 0; x <= strVariantsArray.Length - 1; x++)
              {
                  for (y = 0; y <= oRxPackageItemCollection.Count - 1; y++)
                  {
                      strTableName = "fvs_tree_IN_" + strVariantsArray[x].Trim() + "_P" + oRxPackageItemCollection.Item(y).RxPackageId + "_TREE_CUTLIST";
                      if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, strTableName))
                      {
                          strFVSTreeTableLinkNameList = strFVSTreeTableLinkNameList + strTableName + ",";
                      }
                  }
              }
              cmbDatasource.Items.Clear();
              cmbDatasource.Items.Add("Tree Sample");
              cmbDatasource.Items.Add("Tree Table");
              //load the list into an array
              if (strFVSTreeTableLinkNameList.Trim().Length > 0)
              {
                  strFVSTreeTableLinkNameList = strFVSTreeTableLinkNameList.Substring(0, strFVSTreeTableLinkNameList.Length - 1);
                  m_strFVSTreeTableLinkNameArray = frmMain.g_oUtils.ConvertListToArray(strFVSTreeTableLinkNameList, ",");
                  for (x = 0; x <= m_strFVSTreeTableLinkNameArray.Length - 1; x++)
                  {
                      cmbDatasource.Items.Add(m_strFVSTreeTableLinkNameArray[x].Trim());
                  }
              }

              
             
              oRxPackageItemCollection.Clear();
              oRxPackageItemCollection = null;
          }
      }
      else
      {
          m_strTempDBFile = frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir, "accdb");
          oDao.CreateMDB(m_strTempDBFile);
          oDao.CreateTableLink(m_strTempDBFile, "treesample", frmMain.g_oEnv.strAppDir + "\\db\\treesample.mdb", "treesample");
          oDao.m_DaoWorkspace.Close();
      }
      oDao = null;
      

      
   
      LoadDefaultSingleRecordValues();
      
      
    }
     /// <summary>
    /// Required designer variable.
    /// </summary>
    private System.ComponentModel.IContainer components = null;

    /// <summary>
    /// Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose( bool disposing )
    {
      if( disposing && ( components != null ) )
      {
        components.Dispose();
      }
      base.Dispose( disposing );
    }

    #region Windows Form Designer generated code

    /// <summary>
    /// Required method for Designer support - do not modify
    /// the contents of this method with the code editor.
    /// </summary>
    private void InitializeComponent()
    {
            this.btnTreeVolBatch = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.progressBarBasic1 = new ProgressBarBasic.ProgressBarBasic();
            this.btnCancel = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnDefaultSingle = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.lblDRYBIOM = new System.Windows.Forms.Label();
            this.label20 = new System.Windows.Forms.Label();
            this.lblDRYBIOT = new System.Windows.Forms.Label();
            this.lblVOLCFNET = new System.Windows.Forms.Label();
            this.label18 = new System.Windows.Forms.Label();
            this.label19 = new System.Windows.Forms.Label();
            this.lblVOLCSGRS = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.lblVOLCFGRS = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.label14 = new System.Windows.Forms.Label();
            this.txtRoughCull = new System.Windows.Forms.TextBox();
            this.label13 = new System.Windows.Forms.Label();
            this.txtCull = new System.Windows.Forms.TextBox();
            this.label12 = new System.Windows.Forms.Label();
            this.txtTreeClCd = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.txtStatusCd = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.txtCR = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.txtActualHt = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.txtHt = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.txtDbh = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.txtSpCd = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.txtInvYr = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.txtVolLocGrp = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtStateCd = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.txtPlot = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtCountyCd = new System.Windows.Forms.TextBox();
            this.btnTreeVolSingle = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.btnEdit = new System.Windows.Forms.Button();
            this.btnLoad = new System.Windows.Forms.Button();
            this.cmbDatasource = new System.Windows.Forms.ComboBox();
            this.btnLinkTableTest = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // btnTreeVolBatch
            // 
            this.btnTreeVolBatch.Location = new System.Drawing.Point(387, 286);
            this.btnTreeVolBatch.Name = "btnTreeVolBatch";
            this.btnTreeVolBatch.Size = new System.Drawing.Size(208, 20);
            this.btnTreeVolBatch.TabIndex = 2;
            this.btnTreeVolBatch.Text = "Batch Calculate Volume And Biomass";
            this.btnTreeVolBatch.UseVisualStyleBackColor = true;
            this.btnTreeVolBatch.Click += new System.EventHandler(this.btnTreeVolBatch_Click);
            // 
            // panel1
            // 
            this.panel1.AutoScroll = true;
            this.panel1.Controls.Add(this.progressBarBasic1);
            this.panel1.Controls.Add(this.btnCancel);
            this.panel1.Controls.Add(this.groupBox2);
            this.panel1.Controls.Add(this.groupBox1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(901, 576);
            this.panel1.TabIndex = 3;
            // 
            // progressBarBasic1
            // 
            this.progressBarBasic1.BackColor = System.Drawing.Color.White;
            this.progressBarBasic1.Font = new System.Drawing.Font("Arial", 10.25F);
            this.progressBarBasic1.ForeColor = System.Drawing.Color.ForestGreen;
            this.progressBarBasic1.Location = new System.Drawing.Point(12, 541);
            this.progressBarBasic1.Name = "progressBarBasic1";
            this.progressBarBasic1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.progressBarBasic1.Size = new System.Drawing.Size(803, 23);
            this.progressBarBasic1.TabIndex = 33;
            this.progressBarBasic1.Text = "progressBarBasic1";
            this.progressBarBasic1.TextStyle = ProgressBarBasic.ProgressBarBasic.TextStyleType.Percentage;
            this.progressBarBasic1.Visible = false;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(820, 542);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(69, 22);
            this.btnCancel.TabIndex = 4;
            this.btnCancel.Text = "Cancel\r\n";
            this.btnCancel.UseVisualStyleBackColor = false;
            this.btnCancel.Visible = false;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnDefaultSingle);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Controls.Add(this.label14);
            this.groupBox2.Controls.Add(this.txtRoughCull);
            this.groupBox2.Controls.Add(this.label13);
            this.groupBox2.Controls.Add(this.txtCull);
            this.groupBox2.Controls.Add(this.label12);
            this.groupBox2.Controls.Add(this.txtTreeClCd);
            this.groupBox2.Controls.Add(this.label11);
            this.groupBox2.Controls.Add(this.txtStatusCd);
            this.groupBox2.Controls.Add(this.label10);
            this.groupBox2.Controls.Add(this.txtCR);
            this.groupBox2.Controls.Add(this.label9);
            this.groupBox2.Controls.Add(this.txtActualHt);
            this.groupBox2.Controls.Add(this.label8);
            this.groupBox2.Controls.Add(this.txtHt);
            this.groupBox2.Controls.Add(this.label7);
            this.groupBox2.Controls.Add(this.txtDbh);
            this.groupBox2.Controls.Add(this.label6);
            this.groupBox2.Controls.Add(this.txtSpCd);
            this.groupBox2.Controls.Add(this.label5);
            this.groupBox2.Controls.Add(this.txtInvYr);
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.txtVolLocGrp);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.txtStateCd);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.txtPlot);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.txtCountyCd);
            this.groupBox2.Controls.Add(this.btnTreeVolSingle);
            this.groupBox2.Location = new System.Drawing.Point(12, 322);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(877, 213);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Tree Volumes And Biomass One Record Test";
            // 
            // btnDefaultSingle
            // 
            this.btnDefaultSingle.Location = new System.Drawing.Point(18, 141);
            this.btnDefaultSingle.Name = "btnDefaultSingle";
            this.btnDefaultSingle.Size = new System.Drawing.Size(126, 39);
            this.btnDefaultSingle.TabIndex = 32;
            this.btnDefaultSingle.Text = "Default Values";
            this.btnDefaultSingle.UseVisualStyleBackColor = true;
            this.btnDefaultSingle.Click += new System.EventHandler(this.btnDefaultSingle_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.lblDRYBIOM);
            this.groupBox3.Controls.Add(this.label20);
            this.groupBox3.Controls.Add(this.lblDRYBIOT);
            this.groupBox3.Controls.Add(this.lblVOLCFNET);
            this.groupBox3.Controls.Add(this.label18);
            this.groupBox3.Controls.Add(this.label19);
            this.groupBox3.Controls.Add(this.lblVOLCSGRS);
            this.groupBox3.Controls.Add(this.label17);
            this.groupBox3.Controls.Add(this.lblVOLCFGRS);
            this.groupBox3.Controls.Add(this.label15);
            this.groupBox3.Location = new System.Drawing.Point(500, 15);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(303, 192);
            this.groupBox3.TabIndex = 31;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Results";
            // 
            // lblDRYBIOM
            // 
            this.lblDRYBIOM.BackColor = System.Drawing.Color.Beige;
            this.lblDRYBIOM.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDRYBIOM.Location = new System.Drawing.Point(154, 91);
            this.lblDRYBIOM.Name = "lblDRYBIOM";
            this.lblDRYBIOM.Size = new System.Drawing.Size(142, 29);
            this.lblDRYBIOM.TabIndex = 20;
            this.lblDRYBIOM.Text = "0";
            this.lblDRYBIOM.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label20
            // 
            this.label20.AutoSize = true;
            this.label20.Location = new System.Drawing.Point(185, 73);
            this.label20.Name = "label20";
            this.label20.Size = new System.Drawing.Size(57, 13);
            this.label20.TabIndex = 19;
            this.label20.Text = "DRYBIOM";
            // 
            // lblDRYBIOT
            // 
            this.lblDRYBIOT.BackColor = System.Drawing.Color.Beige;
            this.lblDRYBIOT.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblDRYBIOT.Location = new System.Drawing.Point(4, 152);
            this.lblDRYBIOT.Name = "lblDRYBIOT";
            this.lblDRYBIOT.Size = new System.Drawing.Size(142, 29);
            this.lblDRYBIOT.TabIndex = 18;
            this.lblDRYBIOT.Text = "0";
            this.lblDRYBIOT.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // lblVOLCFNET
            // 
            this.lblVOLCFNET.BackColor = System.Drawing.Color.Beige;
            this.lblVOLCFNET.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLCFNET.Location = new System.Drawing.Point(4, 91);
            this.lblVOLCFNET.Name = "lblVOLCFNET";
            this.lblVOLCFNET.Size = new System.Drawing.Size(142, 29);
            this.lblVOLCFNET.TabIndex = 16;
            this.lblVOLCFNET.Text = "0";
            this.lblVOLCFNET.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(41, 73);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(63, 13);
            this.label18.TabIndex = 15;
            this.label18.Text = "VOLCFNET";
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(49, 130);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(55, 13);
            this.label19.TabIndex = 17;
            this.label19.Text = "DRYBIOT";
            // 
            // lblVOLCSGRS
            // 
            this.lblVOLCSGRS.BackColor = System.Drawing.Color.Beige;
            this.lblVOLCSGRS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLCSGRS.Location = new System.Drawing.Point(154, 38);
            this.lblVOLCSGRS.Name = "lblVOLCSGRS";
            this.lblVOLCSGRS.Size = new System.Drawing.Size(142, 29);
            this.lblVOLCSGRS.TabIndex = 12;
            this.lblVOLCSGRS.Text = "0";
            this.lblVOLCSGRS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(185, 17);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(65, 13);
            this.label17.TabIndex = 11;
            this.label17.Text = "VOLCSGRS";
            // 
            // lblVOLCFGRS
            // 
            this.lblVOLCFGRS.BackColor = System.Drawing.Color.Beige;
            this.lblVOLCFGRS.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblVOLCFGRS.Location = new System.Drawing.Point(4, 38);
            this.lblVOLCFGRS.Name = "lblVOLCFGRS";
            this.lblVOLCFGRS.Size = new System.Drawing.Size(142, 29);
            this.lblVOLCFGRS.TabIndex = 10;
            this.lblVOLCFGRS.Text = "0";
            this.lblVOLCFGRS.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Location = new System.Drawing.Point(41, 17);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(64, 13);
            this.label15.TabIndex = 9;
            this.label15.Text = "VOLCFGRS";
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(363, 77);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(56, 13);
            this.label14.TabIndex = 30;
            this.label14.Text = "RoughCull";
            // 
            // txtRoughCull
            // 
            this.txtRoughCull.Location = new System.Drawing.Point(425, 74);
            this.txtRoughCull.Name = "txtRoughCull";
            this.txtRoughCull.Size = new System.Drawing.Size(62, 20);
            this.txtRoughCull.TabIndex = 29;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(255, 77);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(24, 13);
            this.label13.TabIndex = 28;
            this.label13.Text = "Cull";
            // 
            // txtCull
            // 
            this.txtCull.Location = new System.Drawing.Point(286, 77);
            this.txtCull.Name = "txtCull";
            this.txtCull.Size = new System.Drawing.Size(71, 20);
            this.txtCull.TabIndex = 27;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(131, 80);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(51, 13);
            this.label12.TabIndex = 26;
            this.label12.Text = "TreeClCd";
            // 
            // txtTreeClCd
            // 
            this.txtTreeClCd.Location = new System.Drawing.Point(190, 77);
            this.txtTreeClCd.Name = "txtTreeClCd";
            this.txtTreeClCd.Size = new System.Drawing.Size(59, 20);
            this.txtTreeClCd.TabIndex = 25;
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(15, 80);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(50, 13);
            this.label11.TabIndex = 24;
            this.label11.Text = "StatusCd";
            // 
            // txtStatusCd
            // 
            this.txtStatusCd.Location = new System.Drawing.Point(66, 77);
            this.txtStatusCd.Name = "txtStatusCd";
            this.txtStatusCd.Size = new System.Drawing.Size(59, 20);
            this.txtStatusCd.TabIndex = 23;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(15, 105);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(89, 13);
            this.label10.TabIndex = 22;
            this.label10.Text = "Crown Ratio (CR)";
            // 
            // txtCR
            // 
            this.txtCR.Location = new System.Drawing.Point(104, 102);
            this.txtCR.Name = "txtCR";
            this.txtCR.Size = new System.Drawing.Size(80, 20);
            this.txtCR.TabIndex = 21;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(363, 57);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(48, 13);
            this.label9.TabIndex = 20;
            this.label9.Text = "ActualHt";
            // 
            // txtActualHt
            // 
            this.txtActualHt.Location = new System.Drawing.Point(425, 52);
            this.txtActualHt.Name = "txtActualHt";
            this.txtActualHt.Size = new System.Drawing.Size(62, 20);
            this.txtActualHt.TabIndex = 19;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(255, 57);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(18, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Ht";
            // 
            // txtHt
            // 
            this.txtHt.Location = new System.Drawing.Point(286, 54);
            this.txtHt.Name = "txtHt";
            this.txtHt.Size = new System.Drawing.Size(71, 20);
            this.txtHt.TabIndex = 17;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(131, 57);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(30, 13);
            this.label7.TabIndex = 16;
            this.label7.Text = "DBH";
            // 
            // txtDbh
            // 
            this.txtDbh.Location = new System.Drawing.Point(190, 54);
            this.txtDbh.Name = "txtDbh";
            this.txtDbh.Size = new System.Drawing.Size(59, 20);
            this.txtDbh.TabIndex = 15;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(15, 57);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(33, 13);
            this.label6.TabIndex = 14;
            this.label6.Text = "SpCd";
            // 
            // txtSpCd
            // 
            this.txtSpCd.Location = new System.Drawing.Point(66, 54);
            this.txtSpCd.Name = "txtSpCd";
            this.txtSpCd.Size = new System.Drawing.Size(59, 20);
            this.txtSpCd.TabIndex = 13;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(363, 36);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(32, 13);
            this.label5.TabIndex = 12;
            this.label5.Text = "InvYr";
            // 
            // txtInvYr
            // 
            this.txtInvYr.Location = new System.Drawing.Point(425, 33);
            this.txtInvYr.Name = "txtInvYr";
            this.txtInvYr.Size = new System.Drawing.Size(62, 20);
            this.txtInvYr.TabIndex = 11;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(190, 105);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 13);
            this.label4.TabIndex = 10;
            this.label4.Text = "Vol_Loc_Grp";
            // 
            // txtVolLocGrp
            // 
            this.txtVolLocGrp.Location = new System.Drawing.Point(265, 102);
            this.txtVolLocGrp.Name = "txtVolLocGrp";
            this.txtVolLocGrp.Size = new System.Drawing.Size(80, 20);
            this.txtVolLocGrp.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 36);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "StateCd";
            // 
            // txtStateCd
            // 
            this.txtStateCd.Location = new System.Drawing.Point(66, 33);
            this.txtStateCd.Name = "txtStateCd";
            this.txtStateCd.Size = new System.Drawing.Size(59, 20);
            this.txtStateCd.TabIndex = 7;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(255, 36);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(25, 13);
            this.label2.TabIndex = 6;
            this.label2.Text = "Plot";
            // 
            // txtPlot
            // 
            this.txtPlot.Location = new System.Drawing.Point(286, 33);
            this.txtPlot.Name = "txtPlot";
            this.txtPlot.Size = new System.Drawing.Size(71, 20);
            this.txtPlot.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(131, 36);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 4;
            this.label1.Text = "CountyCd";
            // 
            // txtCountyCd
            // 
            this.txtCountyCd.Location = new System.Drawing.Point(190, 33);
            this.txtCountyCd.Name = "txtCountyCd";
            this.txtCountyCd.Size = new System.Drawing.Size(59, 20);
            this.txtCountyCd.TabIndex = 3;
            // 
            // btnTreeVolSingle
            // 
            this.btnTreeVolSingle.Location = new System.Drawing.Point(193, 138);
            this.btnTreeVolSingle.Name = "btnTreeVolSingle";
            this.btnTreeVolSingle.Size = new System.Drawing.Size(98, 44);
            this.btnTreeVolSingle.TabIndex = 2;
            this.btnTreeVolSingle.Text = "Calculate Volume And Biomass";
            this.btnTreeVolSingle.UseVisualStyleBackColor = true;
            this.btnTreeVolSingle.Click += new System.EventHandler(this.btnTreeVolSingle_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.btnEdit);
            this.groupBox1.Controls.Add(this.btnLoad);
            this.groupBox1.Controls.Add(this.cmbDatasource);
            this.groupBox1.Controls.Add(this.btnLinkTableTest);
            this.groupBox1.Controls.Add(this.btnTreeVolBatch);
            this.groupBox1.Location = new System.Drawing.Point(3, 3);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(886, 313);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Tree Volumes And Biomass Batch Test";
            // 
            // btnEdit
            // 
            this.btnEdit.Location = new System.Drawing.Point(267, 286);
            this.btnEdit.Name = "btnEdit";
            this.btnEdit.Size = new System.Drawing.Size(114, 20);
            this.btnEdit.TabIndex = 7;
            this.btnEdit.Text = "Edit Selected Row";
            this.btnEdit.UseVisualStyleBackColor = true;
            this.btnEdit.Click += new System.EventHandler(this.btnEdit_Click);
            // 
            // btnLoad
            // 
            this.btnLoad.Location = new System.Drawing.Point(184, 285);
            this.btnLoad.Name = "btnLoad";
            this.btnLoad.Size = new System.Drawing.Size(77, 21);
            this.btnLoad.TabIndex = 6;
            this.btnLoad.Text = "Load";
            this.btnLoad.UseVisualStyleBackColor = true;
            this.btnLoad.Click += new System.EventHandler(this.btnLoad_Click);
            // 
            // cmbDatasource
            // 
            this.cmbDatasource.FormattingEnabled = true;
            this.cmbDatasource.Items.AddRange(new object[] {
            "Tree Sample"});
            this.cmbDatasource.Location = new System.Drawing.Point(12, 286);
            this.cmbDatasource.Name = "cmbDatasource";
            this.cmbDatasource.Size = new System.Drawing.Size(166, 21);
            this.cmbDatasource.TabIndex = 5;
            this.cmbDatasource.Text = "Tree Sample";
            // 
            // btnLinkTableTest
            // 
            this.btnLinkTableTest.Location = new System.Drawing.Point(601, 286);
            this.btnLinkTableTest.Name = "btnLinkTableTest";
            this.btnLinkTableTest.Size = new System.Drawing.Size(249, 20);
            this.btnLinkTableTest.TabIndex = 3;
            this.btnLinkTableTest.Text = "Test Link To Oracle BIOSUM_VOLUME table";
            this.btnLinkTableTest.UseVisualStyleBackColor = true;
            this.btnLinkTableTest.Click += new System.EventHandler(this.btnLinkTableTest_Click);
            // 
            // frmFCSTreeVolumeEdit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 576);
            this.Controls.Add(this.panel1);
            this.Name = "frmFCSTreeVolumeEdit";
            this.Text = "Tree Volume and Biomass Calculator Troubleshooter";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmFCSTreeVolumeEdit_FormClosing);
            this.Resize += new System.EventHandler(this.frmFCSTreeVolumeEdit_Resize);
            this.panel1.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.ResumeLayout(false);

    }

    #endregion

       
  
    private void LoadDefaultSingleRecordValues()
    {
        this.txtActualHt.Text = "80";
        this.txtCountyCd.Text = "87";
        this.txtCR.Text = "0";
        this.txtCull.Text = "0";
        this.txtDbh.Text = "20.0";
        this.txtHt.Text = "80";
        this.txtInvYr.Text = "2012";
        this.txtPlot.Text = "3633";
        this.txtRoughCull.Text = "0";
        this.txtStateCd.Text = "6";
        this.txtSpCd.Text = "202";
        this.txtStatusCd.Text = "1";
        this.txtTreeClCd.Text = "2";
        this.txtVolLocGrp.Text = "S26LCA";

     
    }

    private void btnTreeVolSingle_Click( object sender, EventArgs e )
    {
        lblDRYBIOM.Text = "0"; lblDRYBIOT.Text = "0"; lblVOLCFGRS.Text = "0";
        lblVOLCFNET.Text = "0"; lblVOLCSGRS.Text = "0";

      FIADB.Oracle.Services m_oOracleServices = new FIADB.Oracle.Services();
      m_oOracleServices.Start();
      m_oOracleServices.m_oTree.GetVolumesMode = FIADB.Oracle.Services.Tree.GetVolumesModeValues.InsertRowTrigger;
      if( m_oOracleServices.m_intError == 0 )
      {
          try
          {

              m_oOracleServices.m_oTree.InstantiateNewBiosumTreeInputRecord();
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.ActualHt = Convert.ToInt32(txtActualHt.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CND_CN = "1";
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CountyCd = Convert.ToInt32(txtCountyCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.CR = Convert.ToInt32(txtCR.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Cull = Convert.ToInt32(txtCull.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.DBH = Convert.ToDouble(txtDbh.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Ht = Convert.ToInt32(txtHt.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.InvYr = Convert.ToInt32(txtInvYr.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Plot = Convert.ToInt32(txtPlot.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.PLT_CN = "1";
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.RecordId = 1;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.RoughCull = Convert.ToInt32(txtRoughCull.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.SpCd = Convert.ToInt32(txtSpCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.StateCd = Convert.ToInt32(txtStateCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.StatusCd = Convert.ToInt32(txtStatusCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.TRE_CN = "1";
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.TreeClCd = Convert.ToInt32(txtTreeClCd.Text.Trim());
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Tree = 123456;
              m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord.Vol_Loc_Grp = txtVolLocGrp.Text.Trim();

              m_oOracleServices.m_oTree.AddBiosumRecord(m_oOracleServices.m_oTree.BiosumTreeInputSingleRecord);
              if (m_oOracleServices.m_intError == 0)
              {
                  m_oOracleServices.m_oTree.GetBiosumVolumes();
                  if (m_oOracleServices.m_intError == 0)
                  {
                      lblDRYBIOM.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).DRYBIOM.ToString().Trim();
                      lblDRYBIOT.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).DRYBIOT.ToString().Trim();
                      lblVOLCFGRS.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLCFGRS.ToString().Trim();
                      lblVOLCSGRS.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLCSGRS.ToString().Trim();
                      lblVOLCFNET.Text = m_oOracleServices.m_oTree.BiosumTreeInputRecordCollection.Item(0).VOLCFNET.ToString().Trim();

                  }
                  else
                  {
                      lblDRYBIOM.Text = "ERROR";
                      lblDRYBIOT.Text = "ERROR";
                      lblVOLCFGRS.Text = "ERROR";
                      lblVOLCSGRS.Text = "ERROR";
                      lblVOLCFNET.Text = "ERROR";
                      MessageBox.Show(m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
                  }
              }
              else
              {
                  lblDRYBIOM.Text = "ERROR";
                  lblDRYBIOT.Text = "ERROR";
                  lblVOLCFGRS.Text = "ERROR";
                  lblVOLCSGRS.Text = "ERROR";
                  lblVOLCFNET.Text = "ERROR";
                  MessageBox.Show(m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
              }
          }
          catch (Exception err)
          {
              MessageBox.Show(err.Message, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error);
          }
      }
      else
        MessageBox.Show( m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, MessageBoxIcon.Error );

     

    }

    private void btnDefaultSingle_Click(object sender, EventArgs e)
    {
        LoadDefaultSingleRecordValues();
    }

    private void btnTreeVolBatch_Click(object sender, EventArgs e)
    {
        if (m_strGridTableSource.Trim().Length == 0) return;
        RunBatch_Start();
       
    }
    private void RunBatch_Start()
    {
        btnCancel.Visible = true;
        btnCancel.Invalidate();
        btnCancel.Refresh();
        groupBox1.Enabled = false;
        groupBox2.Enabled = false;
        groupBox3.Enabled = false;
        frmMain.g_oDelegate.InitializeThreadEvents();
        frmMain.g_oDelegate.m_oEventStopThread.Reset();
        frmMain.g_oDelegate.m_oEventThreadStopped.Reset();
        frmMain.g_oDelegate.m_oThread = new Thread(new ThreadStart(this.RunBatch_Main));
        frmMain.g_oDelegate.m_oThread.IsBackground = true;
        frmMain.g_oDelegate.m_oThread.Start();
    }
    private void RunBatch_Main()
    {
        string strWorkDbFile = "";
        string strConn = "";
        string strColumns = "";
        string strValues = "";
        int x = 0;
        int y = 0;
        System.Windows.Forms.CurrencyManager oCm;
        System.Data.DataView oDv;
        int intCurrRow = 0;
        int intRecordCount = 0;
        int intThermValue = 0;
        string strTable = "";
       
        dao_data_access oDao = new dao_data_access();
        System.Data.DataRow p_rowFound;

        frmMain.g_oDelegate.CurrentThreadProcessName = "main";
        frmMain.g_oDelegate.CurrentThreadProcessStarted = true;
        
        strTable = m_strGridTableSource;
        
            intRecordCount = Convert.ToInt32(m_oAdo.getRecordCount(
                m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
                "SELECT COUNT(*) FROM " + strTable,
                strTable));

            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Minimum", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", 0);
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Maximum", 100);
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", true);
          
            

            //step 1 - delete old existing link
            if (oDao.TableExists(m_strTempDBFile, "fcs_biosum_volume"))
            {

                oDao.DeleteTableFromMDB(m_strTempDBFile, "fcs_biosum_volume");



            }
            intThermValue++;
           
            UpdateThermPercent(0,intRecordCount+8,intThermValue);
            //step 2 - create new link
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Create Access Table Link To BIOSUM_VOLUME Oracle Table...Stand By");
            strWorkDbFile = m_strTempDBFile;
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            oDao.CreateOracleXETableLink("FIA Biosum Oracle Services", "FCS", "fcs", "FCS", "BIOSUM_VOLUME", strWorkDbFile, "fcs_biosum_volume");
            oDao.m_DaoWorkspace.Close();
            oDao = null;
            
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            //step 3 - open ado connection
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            //step 4 - start oracle XE services
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent,1,"Starting Oracle Services...Stand By");
            FIADB.Oracle.Services m_oOracleServices = new FIADB.Oracle.Services();
            m_oOracleServices.Start();
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);

            if (m_oOracleServices.m_oTree == null) MessageBox.Show("m_oTree==null");
            m_oOracleServices.m_oTree.GetVolumesMode = FIADB.Oracle.Services.Tree.GetVolumesModeValues.SQLUpdate;
            if (m_strGridTableSource.Trim() != Tables.FVS.DefaultOracleInputVolumesTable)
            {
                //step 5 - delete and create work tables
                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputVolumesTable))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.FVS.DefaultOracleInputVolumesTable);
                frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputVolumesTable);

                if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputFCSVolumesTable))
                    m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.FVS.DefaultOracleInputFCSVolumesTable);
                frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputFCSVolumesTable);

                intThermValue++;
                UpdateThermPercent(0, intRecordCount + 8, intThermValue);

                //step 6 - insert records
                frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Prepare Tree Data For Oracle...Stand By");
                strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                            "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,TRE_CN,CND_CN,PLT_CN";


                strValues = "CINT(MID(BIOSUM_COND_ID,6,2)) AS STATECD," +
                            "CINT(MID(BIOSUM_COND_ID,12,3)) AS COUNTYCD," +
                            "CINT(MID(BIOSUM_COND_ID,16,5)) AS PLOT," +
                            "INVYR,VOL_LOC_GRP,ID AS TREE,SPCD,DBH AS DIA,HT,ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL," +
                            "CSTR(ID) AS TRE_CN," +
                            "BIOSUM_COND_ID AS CND_CN," +
                            "MID(BIOSUM_COND_ID,1,LEN(BIOSUM_COND_ID)-1) AS PLT_CN";

                m_oAdo.m_strSQL = "INSERT INTO " + Tables.FVS.DefaultOracleInputFCSVolumesTable + " " +
                                 "(" + strColumns + ") SELECT " + strValues + " FROM " + strTable;
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            }
            else
            {
                strColumns = "STATECD,COUNTYCD,PLOT,INVYR,VOL_LOC_GRP,TREE,SPCD,DIA,HT," +
                            "ACTUALHT,CR,STATUSCD,TREECLCD,ROUGHCULL,CULL,TRE_CN,CND_CN,PLT_CN";
                intThermValue++;
                UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            }

            m_oAdo.m_strSQL = "INSERT INTO fcs_biosum_volume (" + strColumns + ") SELECT " + strColumns + " FROM " + Tables.FVS.DefaultOracleInputFCSVolumesTable;

            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);

            //step 7 - Get returned results from FCS Oracle 
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Wait For Oracle Volume Compilation To Complete...Stand By");
            m_oOracleServices.m_oTree.GetBiosumVolumes();
            intThermValue++;
            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
            //step 8 - Update grid with returned results
            frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Update Grid With Volume Values...Stand By");
            if (m_oOracleServices.m_intError == 0)
            {

                strConn = m_oAdo.m_OleDbConnection.ConnectionString;
                m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);

                m_oAdo.OpenConnection(strConn);

                oCm = (CurrencyManager)this.BindingContext[uc_gridview1.m_dg.DataSource, uc_gridview1.m_dg.DataMember];
                oDv = (DataView)oCm.List;
                intCurrRow = uc_gridview1.m_intCurrRow - 1;
                y = oDv.Count;



                m_oAdo.SqlQueryReader(m_oAdo.m_OleDbConnection, "SELECT * FROM fcs_biosum_volume");
                if (m_oAdo.m_OleDbDataReader.HasRows)
                {
                    DataColumn[] colPk = new DataColumn[1];
                    colPk[0] = uc_gridview1.m_ds.Tables[0].Columns["id"];
                    uc_gridview1.m_ds.Tables[0].PrimaryKey = colPk;
                    while (m_oAdo.m_OleDbDataReader.Read())
                    {
                        if (intThermValue < intRecordCount+8)
                        {
                            intThermValue++;
                            UpdateThermPercent(0, intRecordCount + 8, intThermValue);
                        }
                        System.Object[] p_searchID = new Object[1];
                        p_searchID[0] = Convert.ToInt32(m_oAdo.m_OleDbDataReader["tre_cn"]);
                        p_rowFound = uc_gridview1.m_ds.Tables[0].Rows.Find(p_searchID[0]);
                        if (p_rowFound != null)
                        {
                            if (m_oAdo.m_OleDbDataReader["VOLCSGRS_CALC"] != DBNull.Value)
                                p_rowFound["volcsgrs"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCSGRS_CALC"]);
                            else
                                p_rowFound["volcsgrs"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFGRS_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFGRS"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFGRS_CALC"]);
                            else
                                p_rowFound["VOLCFGRS"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["VOLCFNET_CALC"] != DBNull.Value)
                                p_rowFound["VOLCFNET"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["VOLCFNET_CALC"]);
                            else
                                p_rowFound["VOLCFNET"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIOM_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIOM"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIOM_CALC"]);
                            else
                                p_rowFound["DRYBIOM"] = DBNull.Value;

                            if (m_oAdo.m_OleDbDataReader["DRYBIOT_CALC"] != DBNull.Value)
                                p_rowFound["DRYBIOT"] = Convert.ToDouble(m_oAdo.m_OleDbDataReader["DRYBIOT_CALC"]);
                            else
                                p_rowFound["DRYBIOT"] = DBNull.Value;
                        }
                    }
                }
                m_oAdo.m_OleDbDataReader.Close();
                UpdateThermPercent(0, intRecordCount + 8, intRecordCount + 8);
                System.Threading.Thread.Sleep(2000);
            }
            else
            {
                MessageBox.Show(m_oOracleServices.m_strError, "FIA Biosum", MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
            }
            frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", false);
            frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);
        

        RunBatch_Finished();

       
        frmMain.g_oDelegate.CurrentThreadProcessDone = true;
        frmMain.g_oDelegate.m_oEventThreadStopped.Set();
        this.Invoke(frmMain.g_oDelegate.m_oDelegateThreadFinished);

    }
    private void RunBatch_Finished()
    {
        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Visible", false);
        frmMain.g_oDelegate.SetControlPropertyValue(btnCancel, "Visible", false);
        frmMain.g_oDelegate.SetControlPropertyValue(groupBox1, "Enabled", true);
        frmMain.g_oDelegate.SetControlPropertyValue(groupBox2, "Enabled", true);
        frmMain.g_oDelegate.SetControlPropertyValue(groupBox3, "Enabled", true);
        frmMain.g_oDelegate.SetStatusBarPanelTextValue(frmMain.g_sbpInfo.Parent, 1, "Ready");
    }
    private void UpdateThermPercent(int p_intMin, int p_intMax, int p_intValue)
    {
        int intPercent = (int)(((double)(p_intValue - p_intMin) /
            (double)(p_intMax - p_intMin)) * 100);

        frmMain.g_oDelegate.SetControlPropertyValue(progressBarBasic1, "Value", intPercent);

    }
    private void button1_Click(object sender, EventArgs e)
    {
    }
       

    private void btnLoad_Click(object sender, EventArgs e)
    {
        if (cmbDatasource.Text.Trim().ToUpper() == "TREE SAMPLE") LoadTreeSample();
        else if (cmbDatasource.Text.Trim().ToUpper() == "TREE TABLE")
        {
            FIA_Biosum_Manager.frmDialog oDlg = new frmDialog();
           
            oDlg.uc_select_list_item1.lblTitle.Text = "FVS Variant";
            oDlg.uc_select_list_item1.listBox1.Sorted = true;
            oDlg.uc_select_list_item1.lblMsg.Hide();
            if (m_oAdo.m_OleDbConnection == null)
                m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));

            oDlg.uc_select_list_item1.loadvalues(m_oAdo, m_oAdo.m_OleDbConnection,
                                            "SELECT DISTINCT fvs_variant FROM " + m_oQueries.m_oFIAPlot.m_strPlotTable, "fvs_variant");
            oDlg.uc_select_list_item1.Show();

            DialogResult result = oDlg.ShowDialog();
            if (result == DialogResult.OK)
            {
                if (oDlg.uc_select_list_item1.listBox1.SelectedItems.Count > 0)
                {
                    LoadTreeTable(oDlg.uc_select_list_item1.listBox1.SelectedItems[0].ToString().Trim());
                }
            }
            oDlg.Dispose();
            oDlg = null;
        }
        else
        {
           
            LoadFVSOutTrees(cmbDatasource.Text.Trim().Substring(16,3));
            
        }
    }
    private void LoadFVSOutTrees(string p_strRxPackage)
    {
        frmMain.g_oFrmMain.ActivateStandByAnimation(
            this.WindowState,
            this.Left,
            this.Height,
            this.Width,
            this.Top);
        frmMain.g_sbpInfo.Text = "Loading FVS Out Tree Table Data...Stand By";

        if (m_oAdo.m_OleDbConnection == null)
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));


        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputVolumesTable))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.FVS.DefaultOracleInputVolumesTable);

        frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputVolumesTable);

        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputFCSVolumesTable))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE " + Tables.FVS.DefaultOracleInputFCSVolumesTable);

        frmMain.g_oTables.m_oFvs.CreateOracleInputFCSBiosumVolumesTable(m_oAdo, m_oAdo.m_OleDbConnection, Tables.FVS.DefaultOracleInputFCSVolumesTable);



        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "cull_work_table"))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE cull_work_table");

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.FVSOut_BuiltInputTableForVolumeCalculation_Step1(
            Tables.FVS.DefaultOracleInputVolumesTable, cmbDatasource.Text.Trim(),p_strRxPackage);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.FVSOut_BuildInputTableForVolumeCalculation_Step2(
            Tables.FVS.DefaultOracleInputVolumesTable,
            m_oQueries.m_oFIAPlot.m_strTreeTable,
            m_oQueries.m_oFIAPlot.m_strPlotTable,
            m_oQueries.m_oFIAPlot.m_strCondTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.FVSOut_BuildInputTableForVolumeCalculation_Step3(
            Tables.FVS.DefaultOracleInputVolumesTable,
            m_oQueries.m_oFIAPlot.m_strCondTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.FVSOut_BuildInputTableForVolumeCalculation_Step4(
            "cull_work_table",
            Tables.FVS.DefaultOracleInputVolumesTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.PNWRS.FVSOut_BuildInputTableForVolumeCalculation_Step5(
            "cull_work_table",
            Tables.FVS.DefaultOracleInputVolumesTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.PNWRS.FVSOut_BuildInputTableForVolumeCalculation_Step6(
            "cull_work_table",
            Tables.FVS.DefaultOracleInputVolumesTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

        m_oAdo.m_strSQL = Queries.FVS.VolumesAndBiomass.FVSOut_BuildInputTableForVolumeCalculation_Step7(
            Tables.FVS.DefaultOracleInputVolumesTable,
            Tables.FVS.DefaultOracleInputFCSVolumesTable);
        m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);


        uc_gridview1.LoadGridView(
            m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
            "SELECT VOLCFGRS," +
                   "VOLCSGRS," +
                   "VOLCFNET," +
                   "DRYBIOM," +
                   "DRYBIOT," +
                   "id," +
                   "MID(biosum_cond_id, 6, 2 ) AS state," +
                   "MID(biosum_cond_id,12,3) AS county," +
                   "MID(biosum_cond_id,15,7) AS plot," + 
                   "fvs_variant," +
                   "InvYr," +
                   "SpCd," +
                   "Dbh," +
                   "ROUND(Ht,0) AS Ht," +
                   "vol_loc_grp," +
                   "IIF(actualht IS NULL,Ht,ROUND(actualht,0)) AS actualht," +
                   "statuscd," +
                   "treeclcd," +
                   "cr," +
                   "cull," +
                   "IIF(roughcull IS NULL,0,roughcull) AS roughcull," +
                   "fvs_tree_id," +
                   "biosum_cond_id " +
             "FROM " + Tables.FVS.DefaultOracleInputVolumesTable, this.cmbDatasource.Text.Trim());

        m_strGridTableSource = Tables.FVS.DefaultOracleInputVolumesTable;
        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        frmMain.g_sbpInfo.Text = "Ready";




    }
    private void LoadTreeSample()
    {
        if (m_oAdo.m_OleDbConnection != null)
            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
        uc_gridview1.LoadGridView(
            m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
            "SELECT VOLCFGRS," +
                   "VOLCSGRS," +
                   "VOLCFNET," +
                   "DRYBIOM," +
                   "DRYBIOT," +
                   "id," +
                   "MID(biosum_cond_id, 6, 2 ) AS state," + 
                   "MID(biosum_cond_id,12,3) AS county," + 
                   "MID(biosum_cond_id,15,7) AS plot," + 
                   "'CA' AS fvs_variant," + 
                   "InvYr," +
                   "SpCd," +
                   "Dbh," +
                   "Ht," +
                   "vol_loc_grp," +
                   "IIF(actualht IS NULL,Ht,actualht) AS actualht," +
                   "statuscd," +
                   "treeclcd," +
                   "cr," +
                   "cull," +
                   "IIF(roughcull IS NULL,0,roughcull) AS roughcull," +
                   "fvs_tree_id," +
                   "biosum_cond_id " +
             "FROM TreeSample", "TreeSample");

        m_strGridTableSource = "TreeSample";
    }

    private void LoadTreeTable(string p_strFVSVariant)
    {
        frmMain.g_oFrmMain.ActivateStandByAnimation(
          this.WindowState,
          this.Left,
          this.Height,
          this.Width,
          this.Top);

        if (m_oAdo.m_OleDbConnection == null)
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));

        //
        //CREATE TREE TABLE WORK TABLE
        //
        if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "tree_work_table"))
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection,"DROP TABLE tree_work_table");

            frmMain.g_sbpInfo.Text = "Loading Tree Table Data...Stand By";
            m_oAdo.m_strSQL = frmMain.g_oTables.m_oFvs.CreateOracleInputBiosumVolumesTableSQL("tree_work_table");
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.AddAutoNumber(m_oAdo.m_OleDbConnection, "tree_work_table", "id");
            //
            //POPULATE TREE WORK TABLE
            //
            m_oAdo.m_strSQL = "INSERT INTO tree_work_table " + 
                              "SELECT t.biosum_cond_id," + 
                                     "IIF(p.InvYr IS NULL AND p.MeasYear IS NOT NULL,p.MeasYear,IIF(p.InvYr IS NOT NULL,p.InvYr,null)) AS InvYr," + 
                                     "p.fvs_variant," + 
                                     "t.spcd, t.dia AS dbh," + 
                                     "t.ht, c.vol_loc_grp," + 
                                     "IIF(t.actualht IS NULL,t.Ht,t.actualht) AS actualht," + 
                                     "t.statuscd, t.treeclcd," + 
                                     "IIF(t.cr IS NULL,0,t.cr) AS cr," + 
                                     "IIF(t.cull IS NULL,0,ROUND(t.cull,0)) AS cull," + 
                                     "IIF(t.roughcull IS NULL,0,ROUND(t.roughcull,0)) AS roughcull," + 
                                     "IIF(t.decaycd IS NULL,0,t.decaycd) AS decaycd," + 
                                     "t.fvs_tree_id " + 
                             "FROM " + m_oQueries.m_oFIAPlot.m_strTreeTable + " t," + 
                                       m_oQueries.m_oFIAPlot.m_strCondTable + " c," + 
                                       m_oQueries.m_oFIAPlot.m_strPlotTable + " p " + 
                             "WHERE c.biosum_cond_id = t.biosum_cond_id AND " + 
                                   "p.biosum_plot_id = c.biosum_plot_id AND " + 
                                   "p.fvs_variant='" + p_strFVSVariant + "' AND " + 
                                   "(p.InvYr IS NOT NULL OR p.MeasYear IS NOT NULL)";

            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            //
            //update columns
            //
            //total cull
            //populate treeclcd column
            if (m_oAdo.TableExist(m_oAdo.m_OleDbConnection, "cull_total_work_table") == true)
                m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, "DROP TABLE cull_total_work_table");

            m_oAdo.m_strSQL = "SELECT id, IIF(cull IS NOT NULL AND roughcull IS NOT NULL, cull + roughcull," +
                                           "IIF(cull IS NOT NULL,cull," +
                                           "IIF(roughcull IS NOT NULL, roughcull,0))) AS totalcull " +
                              "INTO cull_total_work_table " +
                              "FROM Tree_Work_Table";

            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "UPDATE Tree_Work_Table a " +
                            "INNER JOIN cull_total_work_table b " +
                            "ON a.id=b.id " +
                            "SET a.treeclcd=" +
                            "IIF(a.SpCd IN (62,65,66,106,133,138,304,321,322,475,756,758,990),3," +
                            "IIF(a.StatusCd=2,3," +
                            "IIF(b.totalcull < 75,2," +
                            "IIF(roughcull > 37.5,3,4))))";
            
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);
            m_oAdo.m_strSQL = "UPDATE tree_work_table  a " +
                                     "INNER JOIN cull_total_work_table b " +
                                     "ON a.id=b.id " +
                                     "SET a.treeclcd=" +
                                     "IIF(a.DecayCd > 1,4,IIF(a.dbh < 9 AND a.SpCd < 300,4,a.treeclcd)) " +
                                     "WHERE a.treeclcd=3 AND a.statuscd=2 AND a.SpCd NOT IN (62,65,66,106,133,138,304,321,322,475,756,758,990)";
            
            m_oAdo.SqlNonQuery(m_oAdo.m_OleDbConnection, m_oAdo.m_strSQL);

            m_oAdo.CloseConnection(m_oAdo.m_OleDbConnection);
            m_oAdo.OpenConnection(m_oAdo.getMDBConnString(m_strTempDBFile, "", ""));


        


        

        uc_gridview1.LoadGridView(
            m_oAdo.getMDBConnString(m_strTempDBFile, "", ""),
            "SELECT VOLCFGRS," +
                   "VOLCSGRS," +
                   "VOLCFNET," +
                   "DRYBIOM," +
                   "DRYBIOT," +
                   "id," +
                   "MID(biosum_cond_id, 6, 2 ) AS state," +
                   "MID(biosum_cond_id,12,3) AS county," +
                   "MID(biosum_cond_id,15,7) AS plot," +
                   "fvs_variant," + 
                   "InvYr," +
                   "SpCd," +
                   "Dbh," +
                   "Ht," +
                   "vol_loc_grp," +
                   "IIF(actualht IS NULL,Ht,actualht) AS actualht," +
                   "statuscd," +
                   "treeclcd," +
                   "cr," +
                   "cull," +
                   "IIF(roughcull IS NULL,0,roughcull) AS roughcull," +
                   "fvs_tree_id," +
                   "biosum_cond_id " +
             "FROM Tree_Work_Table", m_oQueries.m_oFIAPlot.m_strTreeTable);

        m_strGridTableSource = "Tree_Work_Table";
        frmMain.g_oFrmMain.DeactivateStandByAnimation();
        frmMain.g_sbpInfo.Text = "Ready";
    }


    private void btnEdit_Click(object sender, EventArgs e)
    {
        this.txtActualHt.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1,COL_ACTUALHT].ToString().Trim() ;
        this.txtCountyCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_COUNTY].ToString().Trim();
        this.txtCR.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_CR].ToString().Trim();
        this.txtCull.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_CULL].ToString().Trim();
        this.txtDbh.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_DBH].ToString().Trim();
        this.txtHt.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_HT].ToString().Trim();
        this.txtInvYr.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_INVYR].ToString().Trim();
        this.txtPlot.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_PLOT].ToString().Trim();
        this.txtRoughCull.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1,COL_ROUGHCULL].ToString().Trim();
        this.txtStateCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_STATE].ToString().Trim();
        this.txtSpCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_SPCD].ToString().Trim();
        this.txtStatusCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_STATUSCD].ToString().Trim();
        this.txtTreeClCd.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_TREECLCD].ToString().Trim();
        this.txtVolLocGrp.Text = uc_gridview1.m_dg[uc_gridview1.m_intCurrRow - 1, COL_VOLLOCGRP].ToString().Trim();
        
    }

    private void frmFCSTreeVolumeEdit_Resize(object sender, EventArgs e)
    {
        ResizeForm();
    }
    private void ResizeForm()
    {
        //width
        groupBox2.Width = this.ClientSize.Width - (int)(groupBox2.Left * 2);
        groupBox1.Width = groupBox2.Width;
        btnCancel.Left = this.ClientSize.Width - btnCancel.Width - 5;

        progressBarBasic1.Width = btnCancel.Left - progressBarBasic1.Left - 3;
        lblPerc.Left = (int)(progressBarBasic1.Width / 2) - (int)(lblPerc.Width * .5);

        //top
        progressBarBasic1.Top = this.ClientSize.Height - progressBarBasic1.Height - 5;
        btnCancel.Top = progressBarBasic1.Top;
        groupBox2.Top = progressBarBasic1.Top - groupBox2.Height - 5;

        //height
        groupBox1.Height = groupBox2.Top - groupBox1.Top - 5;
        btnEdit.Top = groupBox1.Height - btnEdit.Height - 5;
        cmbDatasource.Top = btnEdit.Top;
        btnLoad.Top = btnEdit.Top;
        btnTreeVolBatch.Top = btnEdit.Top;
        btnLinkTableTest.Top = btnEdit.Top;
        uc_gridview1.Height = cmbDatasource.Top - uc_gridview1.Top - 5;
        uc_gridview1.Width = groupBox1.Width - (uc_gridview1.Left * 2) - 5;

    }

    private void btnCancel_Click(object sender, EventArgs e)
    {
        frmMain.g_oDelegate.CurrentThreadProcessSuspended = true;
        frmMain.g_oDelegate.m_oThread.Suspend();
        bool bAbort = frmMain.g_oDelegate.AbortProcessing("FIA Biosum", "Cancel Running The Volume and Biosum Calculator (Y/N)?");
        if (bAbort)
        {
            
            if (frmMain.g_oDelegate.m_oThread.IsAlive)
            {
                frmMain.g_oDelegate.m_oThread.Join();
            }
            frmMain.g_oDelegate.StopThread();
            RunBatch_Finished();

            frmMain.g_oDelegate.m_oThread = null;

        }
        else frmMain.g_oDelegate.m_oThread.Resume();
    }

    private void frmFCSTreeVolumeEdit_FormClosing(object sender, FormClosingEventArgs e)
    {
        if (frmMain.g_oDelegate.CurrentThreadProcessStarted == true &&
            frmMain.g_oDelegate.CurrentThreadProcessDone == false)
        {
            btnCancel_Click(null, null);
            if (frmMain.g_oDelegate.CurrentThreadProcessAborted == true)
            {
            }
            else
            {
                e.Cancel = true; return;
            }

            

        }
        this.Dispose();
    }

    private void btnLinkTableTest_Click(object sender, EventArgs e)
    {
        string strFile=frmMain.g_oUtils.getRandomFile(frmMain.g_oEnv.strTempDir,"accdb");
        string str = "ODBC\r\n-----------------\r\nData Source Name:FIA Biosum Oracle Services\r\nTNS Service Name:XE\r\nUser Id:fcs\r\nOracle Table Name:BIOSUM_VOLUME\r\nMS Access Table Link Name:fcs_biosum_volume\r\nConnection Status:";
        try
        {
            frmMain.g_oFrmMain.ActivateStandByAnimation(
                this.WindowState,
                this.Left,
                this.Height,
                this.Width,
                this.Top);
            FIA_Biosum_Manager.dao_data_access oDao = new FIA_Biosum_Manager.dao_data_access();
            oDao.CreateMDB(strFile);

            oDao.CreateOracleXETableLink("FIA Biosum Oracle Services", "FCS", "fcs", "FCS", "BIOSUM_VOLUME", strFile, "fcs_biosum_volume");
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            if (oDao.TableExists(strFile, "fcs_biosum_volume"))
            {
                MessageBox.Show(str + "Success");
            }
            else
            {
                MessageBox.Show(str + "Error creating MS Access table link to Oracle XE table 'BIOSUM_VOLUME'", "FIA Biosum");
            }
            oDao.m_DaoWorkspace.Close();
            oDao = null;
            
            
        }
        catch (Exception err)
        {
            frmMain.g_oFrmMain.DeactivateStandByAnimation();
            MessageBox.Show(str + "Error creating MS Access table link to Oracle XE table 'BIOSUM_VOLUME'\r\n Error Message:" + err.Message,"FIA Biosum");
        }
        
    
    }
  }
}