# examly - Statistical Metrics and Reporting Tool
library(shiny)
library(dplyr); library(tidyr); library(purrr); library(stringr)
library(readr); library(readxl)
library(officer); library(flextable); library(glue)
library(magrittr)
library(ggplot2)

app_css <- HTML("
  body { background: #f4f7fb; color: #162033; padding-bottom: 58px; }
  .navbar-default {
    background: rgba(255,255,255,.96);
    border: 0;
    border-bottom: 1px solid #dbe4ef;
    box-shadow: 0 8px 24px rgba(15,23,42,.06);
  }
  .navbar-default .navbar-brand { color: #162033; font-weight: 800; letter-spacing: 0; }
  .navbar-default .navbar-nav > li > a { color: #475569; font-weight: 650; }
  .navbar-default .navbar-nav > .active > a,
  .navbar-default .navbar-nav > .active > a:focus,
  .navbar-default .navbar-nav > .active > a:hover { background: #eaf1ff; color: #1d4ed8; }
  .container-fluid { max-width: 1280px; }
  .app-sidebar, .surface-panel, .table-card, details.toggle-card {
    background: #fff;
    border: 1px solid #dbe4ef;
    border-radius: 8px;
    box-shadow: 0 16px 40px rgba(15,23,42,.08);
  }
  .app-sidebar, .surface-panel, .table-card { padding: 18px; margin-bottom: 18px; }
  .page-head { display: flex; align-items: flex-end; justify-content: space-between; gap: 16px; margin: 16px 0 14px; }
  .page-head h3 { margin: 0; font-size: 24px; font-weight: 800; letter-spacing: 0; }
  .metric-grid { display: grid; grid-template-columns: repeat(4, minmax(150px, 1fr)); gap: 12px; margin: 10px 0 18px; }
  .metric-card {
    background: linear-gradient(180deg, #fff 0%, #f8fafc 100%);
    border: 1px solid #dbe4ef;
    border-radius: 8px;
    padding: 14px;
    min-height: 92px;
    box-shadow: 0 10px 24px rgba(15,23,42,.06);
  }
  .metric-label { color: #64748b; font-size: 12px; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; margin-bottom: 8px; }
  .metric-value { color: #162033; font-size: 26px; line-height: 1.1; font-weight: 800; letter-spacing: 0; }
  .metric-note { margin-top: 6px; color: #64748b; font-size: 12px; }
  .table-card table, .surface-panel table, table.table {
    background: #fff;
    border-collapse: separate;
    border-spacing: 0;
    border: 1px solid #dbe4ef;
    border-radius: 8px;
    overflow: hidden;
    width: 100%;
  }
  .table-card th, .surface-panel th, table.table th {
    background: #eef4ff;
    color: #1e3a8a;
    font-weight: 750;
    border-bottom: 1px solid #dbe4ef !important;
  }
  .table-card td, .surface-panel td, table.table td { border-color: #edf2f7 !important; vertical-align: middle !important; }
  .table-card tbody tr:hover, .surface-panel tbody tr:hover, table.table tbody tr:hover { background: #f8fbff; }
  .preview-scroll { max-width: 100%; overflow-x: auto; padding-bottom: 6px; }
  .preview-scroll table { width: max-content; min-width: 100%; }
  .preview-scroll th, .preview-scroll td { white-space: nowrap; }
  .value-pill {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    min-width: 62px;
    padding: 4px 9px;
    border-radius: 8px;
    color: #fff;
    font-weight: 800;
  }
  .value-pill-neutral { background: #64748b; }
  .value-pill-good { background: #16a34a; }
  .value-pill-risk { background: #dc2626; }
  .comment-table { margin-top: 10px; }
  .comment-table td:first-child { width: 220px; font-weight: 750; color: #1e3a8a; }
  .comment-table .comment-state { width: 110px; text-align: center; }
  .outcome-grid-wrap { max-height: 360px; overflow: auto; border: 1px solid #dbe4ef; border-radius: 8px; }
  .outcome-grid { width: 100%; border-collapse: separate; border-spacing: 0; }
  .outcome-grid th { position: sticky; top: 0; z-index: 1; background: #eef4ff; color: #1e3a8a; }
  .outcome-grid th, .outcome-grid td { padding: 7px 9px; border-bottom: 1px solid #edf2f7; vertical-align: middle; }
  .outcome-grid td:first-child { width: 180px; font-weight: 750; color: #334155; }
  .outcome-grid .form-group { margin-bottom: 0; }
  .failure-rank-title { margin: 18px 0 10px; font-weight: 800; color: #162033; }
  .summary-stat-table { font-size: 15px; }
  .summary-stat-table th { font-size: 15px; }
  .summary-stat-table td { font-size: 15px; padding: 10px 12px !important; }
  .summary-stat-table td:first-child { font-weight: 700; color: #1e3a8a; }
  .item-type-grid {
    display: grid;
    grid-template-columns: repeat(4, minmax(190px, 1fr));
    gap: 12px;
  }
  .item-type-card {
    border: 1px solid #dbe4ef;
    border-radius: 8px;
    background: #fff;
    min-height: 230px;
    padding: 12px;
    box-shadow: 0 10px 24px rgba(15,23,42,.05);
  }
  .item-type-card-head {
    display: flex;
    align-items: center;
    justify-content: space-between;
    gap: 8px;
    margin-bottom: 8px;
  }
  .item-type-card-title {
    font-weight: 800;
    color: #162033;
    line-height: 1.25;
  }
  .item-type-card .checkbox {
    margin-top: 6px;
    margin-bottom: 6px;
  }
  .item-type-card .form-group {
    max-height: 210px;
    overflow-y: auto;
    margin-bottom: 0;
    padding-right: 4px;
  }
  .select-all-link {
    display: inline-flex;
    align-items: center;
    border: 1px solid #bfdbfe;
    border-radius: 999px;
    padding: 3px 8px;
    background: #eff6ff;
    color: #1d4ed8 !important;
    font-size: 12px;
    font-weight: 700;
    text-decoration: none !important;
    white-space: nowrap;
  }
  @media (max-width: 1100px) {
    .item-type-grid { grid-template-columns: repeat(2, minmax(190px, 1fr)); }
  }
  @media (max-width: 700px) {
    .item-type-grid { grid-template-columns: 1fr; }
  }
  .btn, .form-control, .selectize-input { border-radius: 8px !important; }
  .btn-default, .btn-primary, .btn-info, .shiny-download-link {
    background: #2563eb;
    border: 0;
    color: #fff !important;
    box-shadow: 0 8px 18px rgba(37,99,235,.16);
  }
  .btn-default:hover, .btn-primary:hover, .btn-info:hover, .shiny-download-link:hover { background: #1d4ed8; color: #fff !important; }
  .report-grid { display: grid; grid-template-columns: repeat(2, minmax(210px, 1fr)); gap: 12px; max-width: 620px; }
  .report-grid .btn { width: 100%; text-align: left; padding: 12px 14px; }
  @media (max-width: 900px) {
    .metric-grid, .report-grid { grid-template-columns: 1fr; }
    .page-head { align-items: flex-start; flex-direction: column; }
  }
")

title_default <- "Sınav Başlığı"

ui <- navbarPage(
  title = uiOutput("app_brand"),
  windowTitle = "examly: Statistical Metrics and Reporting Tool",
  id = "mainnav",
  header = tagList(
    tags$head(tags$style(app_css)),
    tags$script(HTML("
      Shiny.addCustomMessageHandler('set-title', function(txt){
        document.title = txt;
      });
    "))
  ),

  tabPanel(
    title = textOutput("tab_upload", container = span),
    sidebarLayout(
      sidebarPanel(
        width = 4,
        class = "app-sidebar",
        uiOutput("lang_selector"),
        textInput("exam_title", "Sınav Başlığı", value = title_default),
        textInput("school_name", "Okul", value = ""),
        textInput("class_name", "Sınıf/Şube", value = ""),
        uiOutput("file_input"),
        checkboxInput("header", "Başlık satırı var (CSV için)", TRUE),
        radioButtons(
          "sep", "Ayraç (CSV için)",
          choices = c("Virgül" = ",", "Noktalı Virgül" = ";"),
          selected = ","
        ),

        uiOutput("col_selectors"),
        tags$hr()
      ),
      mainPanel(
        width = 8,
        tags$div(
          class = "surface-panel",
          tags$div(class = "page-head", h3(textOutput("title_display"))),
          tableOutput("dimensions_info"),
          tags$div(class = "preview-scroll", tableOutput("sample_head"))
        )
      )
    )
  ),

  tabPanel(title = textOutput("tab_keys", container = span),
           fluidPage(
             tags$head(tags$style(HTML("
        .center-title {text-align:center;font-weight:600;font-size:20px;}
        .compact-wrap {position:relative; overflow-x:hidden; padding:6px; border:1px solid #eee; border-radius:8px;}
        table.compact-grid {border-collapse:separate; border-spacing:0; width:100%; table-layout:fixed;}
        table.compact-grid th {padding:6px 8px; white-space:normal; word-break:break-word; font-size:12px;}
        table.compact-grid td {padding:2px; vertical-align:middle;}
        .grid-input .form-control {position:relative; z-index:1; width:100%; height:24px; padding:4px 6px;}
        .weights-grid th {font-size:11px;}
        .weights-grid td {padding:1px;}
        .grid-input-num .form-control {position:relative; z-index:1; width:100%; height:24px; padding:2px 4px;}
        .section-gap {margin-top:10px; margin-bottom:8px;}
        .muted {color:#777;}
        details.toggle-card {border:1px solid #e5e7eb; border-radius:10px; background:#f9fafb;}
        details.toggle-card[open] {background:#f3f4f6;}
        details.toggle-card > summary {list-style:none; font-weight:700; font-size:16px; padding:10px 40px 10px 14px;
                                       cursor:pointer; position:relative; user-select:none;}
        details.toggle-card > summary::-webkit-details-marker {display:none;}
        details.toggle-card > summary::after {content:'\\25BE'; position:absolute; right:12px; top:50%; transform:translateY(-50%);}
        details.toggle-card[open] > summary::after {content:'\\25B4';}
        details.toggle-card .toggle-body {padding:6px 14px 12px;}
        details.toggle-card .hint {margin:-6px 14px 8px 14px; color:#6b7280;}
        .banner {margin:10px 0; padding:10px 14px; background:#eef6ff; border-left:4px solid #1d4ed8; border-radius:8px; font-size:16px;}
        .details-summary-note {font-size: 12px; font-weight: 400; color: #6b7280;}
        details.toggle-card > summary { display: flex; align-items: center; gap: 8px;}
        .select-all-link { font-size: 11px; text-decoration: underline; cursor: pointer; color: #337ab7; font-weight: normal;}
      "))),
             div(class="center-title", textOutput("title_display_center")),
             br(),
             tags$details(class="toggle-card", open = NA,
                          tags$summary(
                            tagList(
                              textOutput("set_item_types", container = span),
                              tags$span(class = "details-summary-note",
                                        textOutput("set_item_types_note", container = span))
                            )
                          ),
                          tags$div(class="toggle-body",
                                   tags$div(
                                     class = "item-type-grid",
                                     uiOutput("mc_item_selector"),
                                     uiOutput("tf_item_selector"),
                                     uiOutput("bin_item_selector"),
                                     uiOutput("lc_item_selector")
                                   )
                          )
             ),
             uiOutput("max_score_banner"),
             tags$div(class = "surface-panel", uiOutput("learning_outcomes_ui")),
             br(),
             fluidRow(
               column(12,
                      h4(textOutput("mc_key_header")),
                      selectInput("mc_choice_count", label = NULL,
                                  choices = c("A B C D E"=5,"A B C D"=4,"A B C"=3),
                                  selected = 5, width="220px"),
                      div(class="compact-wrap section-gap", uiOutput("mc_answer_table"))
               )
             ),
             fluidRow(
               column(12,
                      h4(textOutput("tf_key_header")),
                      div(class="compact-wrap section-gap", uiOutput("tf_answer_table"))
               )
             ),
             fluidRow(
               column(12,
                      h4(textOutput("lc_header")),
                      div(class="muted", textOutput("lc_note"))
               )
             ),
             hr(),
             div(style="display:flex; justify-content:space-between; align-items:flex-end;",
                 h4(textOutput("weights_header")),
                 uiOutput("ui_distribute_weights_btn")
             ),
             div(class="compact-wrap section-gap", uiOutput("item_weight_table")),
             br()
           )
  ),

  tabPanel(
    title = textOutput("tab_summary", container = span),
    fluidPage(
      tags$div(class = "page-head", h3(textOutput("title_display3"))),
      uiOutput("summary_cards"),
      tags$div(class = "table-card", uiOutput("test_summary_table")),
      hr(),
      h4(textOutput("what_it_means")),
      uiOutput("overall_comment")
    )
  ),

  tabPanel(
    title = textOutput("tab_itemstats", container = span),
    fluidPage(
      tags$div(class = "page-head", h3(textOutput("title_display4"))),
      tags$div(class = "table-card", tableOutput("item_stats_table")),
      hr(),
      h4(textOutput("what_it_means_items")),
      uiOutput("item_comment_ui")
    )
  ),

  tabPanel(
    title = textOutput("tab_distractor", container = span),
    fluidPage(
      tags$div(class = "page-head", h3(textOutput("title_display8"))),
      tags$div(class = "surface-panel", uiOutput("distractor_item_picker")),
      tags$hr(),
      h4(textOutput("distractor_highlights_header")),
      uiOutput("distractor_highlights_ui"),
      hr(), h4(textOutput("distractor_comment_hdr")),
      tags$div(class = "surface-panel", verbatimTextOutput("distractor_comment")),
      hr(), h4(textOutput("distractor_table_hdr")),
      tags$div(class = "table-card", tableOutput("distractor_table"))
    )
  ),

  tabPanel(
    title = textOutput("tab_student_results", container = span),
    fluidPage(
      tags$div(class = "page-head", h3(textOutput("title_display5"))),
      tags$div(class = "table-card", uiOutput("student_table_html")),
      tags$hr(),
      h4(textOutput("student_summary_header", container = span)),
      tags$div(class = "surface-panel", uiOutput("student_summary_text_ui")),
      tags$div(class = "surface-panel", plotOutput("student_grade_plot", height = "420px")),
      tags$div(class = "table-card", uiOutput("item_failure_rank_ui"))
    )
  ),

  tabPanel(
    title = textOutput("tab_reports", container = span),
    fluidPage(
      tags$div(
        class = "page-head",
        h3(textOutput("title_display_rapor")),
        uiOutput("reports_about_button")
      ),
      tags$div(class = "surface-panel", uiOutput("reports_intro_ui")),
      br(),
      tags$div(class = "surface-panel", uiOutput("reports_buttons_ui"))
    )
  ),
  footer = tagList(
    tags$style(HTML("
    body {
        padding-bottom: 70px;
      }
      .custom-footer {position: fixed; left: 0; right: 0; bottom: 0;
        height: 32px; background: #f8f9fa; border-top: 1px solid #e5e7eb;
        display: flex; align-items: center; justify-content: flex-end;
        padding: 0 14px; font-size: 14px; color: #333; z-index: 1040;}
      .custom-footer a {text-decoration: none; color: #337ab7;}
      .custom-footer a:hover {text-decoration: underline;}
      .custom-footer .glyphicon {margin-right: 6px; font-size: 13px;}
    ")),
    tags$div(class = "custom-footer", uiOutput("about_link")),
    tags$script(HTML("
      $(document).on('click', '#about_footer', function(e){
        e.preventDefault();
        Shiny.setInputValue('about_footer', new Date().getTime());
      });"))
  )
)


server <- function(input, output, session) {

  dict <- reactiveVal(examly::i18n_load("tr"))
  T_   <- function(key, default = NULL) examly::i18n_t(dict(), key, default)

  observeEvent(input$lang, {
    dict(examly::i18n_load(input$lang))
  }, ignoreInit = FALSE)

  output$app_brand <- renderUI({
    dict()
    tags$span(T_("AppTitle", "examly: Statistical Metrics and Reporting Tool"))
  })

  observe({
    dict()
    session$sendCustomMessage(
      "set-title",
      T_("AppTitle", "examly: Statistical Metrics and Reporting Tool")
    )
  })

  output$tab_upload <- renderText( T_("UploadTab", "Veri GiriÃƒâ€¦Ã…Â¸i") )
  output$tab_keys <- renderText( T_("KeysTab", "Madde TÃƒÆ’Ã‚Â¼rleri & Anahtarlar") )
  output$set_item_types      <- renderText( T_("SetItemTypes", "Madde Tiplerini Ayarla") )
  output$set_item_types_note <- renderText( T_("SetItemTypesNote",
                                               "(TÃƒÆ’Ã‚Â¼m maddeler 'ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli' kabul edilir. Madde tiplerini deÃƒâ€Ã…Â¸iÃƒâ€¦Ã…Â¸tirmek iÃƒÆ’Ã‚Â§in tÃƒâ€Ã‚Â±klayÃƒâ€Ã‚Â±n.)") )
  output$tab_summary <- renderText(T_("TestSummary", "Test ÃƒÆ’Ã¢â‚¬â€œzeti"))
  output$summary_colour_scale <- renderText(T_("SummaryColourScale", "Renk SkalasÃƒâ€Ã‚Â±"))
  output$visual_cues          <- renderText(T_("VisualCues", "GÃƒÆ’Ã‚Â¶rsel GÃƒÆ’Ã‚Â¶stergeler"))
  output$what_it_means        <- renderText(T_("WhatItMeans", "Ne anlama geliyor?"))
  output$tab_itemstats      <- renderText({ T_("ItemStats", "Madde Ãƒâ€Ã‚Â°statistikleri") })
  output$what_it_means_items<- renderText({ T_("WhatItMeans", "Yorum ve AralÃƒâ€Ã‚Â±klar") })
  output$tab_distractor         <- renderText( T_("Distractor", "ÃƒÆ’Ã¢â‚¬Â¡eldirici Analizi") )
  output$distractor_comment_hdr <- renderText( T_("DistractorCommentHeader", "Yorum") )
  output$distractor_highlights_header <- renderText({ dict(); T_("DistractorHighlightsHeader", "Madde ÃƒÆ’Ã¢â‚¬â€œzeti") })
  output$distractor_table_hdr   <- renderText( T_("DistractorTableHeader", "SeÃƒÆ’Ã‚Â§enek DaÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±lÃƒâ€Ã‚Â±mlarÃƒâ€Ã‚Â±") )
  output$tab_student_results <- renderText( T_("StudentResults", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SonuÃƒÆ’Ã‚Â§larÃƒâ€Ã‚Â±") )

  output$tab_reports <- renderText( T_("Reports", "Raporlar") )

  output$reports_intro_ui <- renderUI({
    dict()
    tagList(
      p( T_("ReportsIntro", "AÃƒâ€¦Ã…Â¸aÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±daki butonlara tÃƒâ€Ã‚Â±klayarak analiz sonuÃƒÆ’Ã‚Â§larÃƒâ€Ã‚Â±nÃƒâ€Ã‚Â± DOCX olarak indirebilirsiniz.") ),
      p( em( T_("ReportsNote", "(ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci ve madde sayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±na gÃƒÆ’Ã‚Â¶re uzun sÃƒÆ’Ã‚Â¼rebilir. LÃƒÆ’Ã‚Â¼tfen bekleyiniz.)") ) )
    )
  })

  output$reports_buttons_ui <- renderUI({
    dict()
    tags$div(
      class = "report-grid",
      downloadButton("dl_rapor_ogrenci",   T_("DownloadStudentDocx",   "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SonuÃƒÆ’Ã‚Â§larÃƒâ€Ã‚Â± (DOCX)")),
      downloadButton("dl_rapor_madde",     T_("DownloadItemDocx",      "Madde Ãƒâ€Ã‚Â°statistikleri (DOCX)")),
      downloadButton("dl_rapor_celdirici", T_("DownloadDistractorDocx","ÃƒÆ’Ã¢â‚¬Â¡eldirici Analizi (DOCX)")),
      downloadButton("dl_rapor_ozet",      T_("DownloadSummaryDocx",   "Test ÃƒÆ’Ã¢â‚¬â€œzeti Ãƒâ€Ã‚Â°statistikleri (DOCX)"))
    )
  })

  output$low_label  <- renderText(T_("Low", "Düşük"))
  output$high_label <- renderText(T_("High", "Yüksek"))
  output$lang_selector <- renderUI({
    dict()
    radioButtons(
      inputId = "lang",
      label   = T_("LanguageSelection", "Dil Seçimi"),
      choices = c("Türkçe" = "tr", "English" = "en"),
      selected <- examly::`%||%`(input$lang, "tr"),
      inline = TRUE
    )
  })

  output$file_input <- renderUI({
    input$lang; dict()

    fileInput(
      inputId = "file",
      label       = T_("UploadData", "Veri Yükleyin (CSV / Excel)"),
      buttonLabel = T_("ChooseFile", "Dosya Seç"),
      placeholder = T_("NoFile", "Henüz seçilmedi"),
      accept = c(".csv",".xlsx",".xls",
                 "text/csv",
                 "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                 "application/vnd.ms-excel")
    )
  })

  prev_title_default <- reactiveVal(NULL)

  output$mc_key_header <- renderText( T_("MCKeyHeader","ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli Cevap AnahtarÃƒâ€Ã‚Â±") )
  output$tf_key_header <- renderText( T_("TFKeyHeader","DoÃƒâ€Ã…Â¸ru/YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸ Maddeleri") )
  output$lc_header     <- renderText( T_("LCHeader","Uzun CevaplÃƒâ€Ã‚Â± (AÃƒÆ’Ã‚Â§Ãƒâ€Ã‚Â±k UÃƒÆ’Ã‚Â§lu) Maddeler") )
  output$lc_note       <- renderText( T_("LCNote","Not: Uzun cevaplÃƒâ€Ã‚Â± maddelerin veri setindeki deÃƒâ€Ã…Â¸erleri ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸rencinin o soruda aldÃƒâ€Ã‚Â±Ãƒâ€Ã…Â¸Ãƒâ€Ã‚Â± puan olmalÃƒâ€Ã‚Â±dÃƒâ€Ã‚Â±r.") )
  output$weights_header<- renderText( T_("WeightsHeader","Madde Puan KatsayÃƒâ€Ã‚Â±larÃƒâ€Ã‚Â± (Maksimum)") )
  output$tab_student_results <- renderText( T_("StudentResults", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SonuÃƒÆ’Ã‚Â§larÃƒâ€Ã‚Â±") )
  output$student_summary_header <- renderText( T_("StudentSummaryHeader", "Puan DaÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±lÃƒâ€Ã‚Â±m ÃƒÆ’Ã¢â‚¬â€œzeti") )

  observeEvent(dict(), {
    new_default <- T_("ExamTitleDefault", "Sınav Başlığı")
    cur <- isolate(input$exam_title)
    prv <- isolate(prev_title_default())

    if (is.null(cur) || cur == "" || (!is.null(prv) && cur == prv)) {
      updateTextInput(
        session, "exam_title",
        label = T_("ExamTitleLabel", "Sınav Başlığı"),
        value = new_default
      )
    } else {
      updateTextInput(
        session, "exam_title",
        label = T_("ExamTitleLabel", "Sınav Başlığı")
      )
    }
    prev_title_default(new_default)

    updateTextInput(session, "school_name", label = T_("SchoolName", "Okul"))
    updateTextInput(session, "class_name", label = T_("ClassName", "Sınıf/Şube"))

    updateCheckboxInput(
      session, "header",
      label = T_("HasHeader", "Başlık satırı var (CSV için)")
    )

    updateRadioButtons(
      session, "sep",
      label   = T_("Separator", "Ayraç (CSV için)"),
      choices = setNames(
        c(",", ";"),
        c(T_("Comma", "Virgül"), T_("Semicolon", "Noktalı Virgül"))
      ),
      selected <- examly::`%||%`(input$sep, ",")
    )

    updateSelectInput(
      session, "mc_choice_count",
      label = T_("MCChoiceCount", "Seçenek sayısı")
    )
  }, ignoreInit = FALSE)

  exam_title <- reactive({ req(input$exam_title); input$exam_title })
  school_name <- reactive({ stringr::str_squish(examly::`%||%`(input$school_name, "")) })
  class_name <- reactive({ stringr::str_squish(examly::`%||%`(input$class_name, "")) })

  add_report_metadata <- function(doc) {
    school <- school_name()
    class_value <- class_name()
    if (nzchar(school) || nzchar(class_value)) {
      if (nzchar(school)) {
        doc <- body_add_par(doc, paste0(T_("SchoolName", "Okul"), ": ", school), style = "Normal")
      }
      if (nzchar(class_value)) {
        doc <- body_add_par(doc, paste0(T_("ClassName", "Sınıf/Şube"), ": ", class_value), style = "Normal")
      }
      doc <- body_add_par(doc, "")
    }
    doc
  }

  output$title_display  <- renderText({ exam_title() })
  output$title_display_center <- renderText({ exam_title() })
  output$title_display3 <- renderText({ exam_title() })
  output$title_display4 <- renderText({ exam_title() })
  output$title_display5 <- renderText({ exam_title() })
  output$title_display8 <- renderText({ exam_title() })
  output$title_display_rapor <- renderText({ exam_title() })

  app_version <- function() {
    desc_paths <- c(file.path(getwd(), "DESCRIPTION"), file.path(getwd(), "..", "..", "DESCRIPTION"))
    for (desc_path in desc_paths) {
      pkg_ver <- tryCatch(as.character(read.dcf(desc_path, fields = "Version")[1, 1]), error = function(e) NA_character_)
      if (!is.na(pkg_ver) && nzchar(pkg_ver)) return(pkg_ver)
    }
    tryCatch(
      as.character(utils::packageDescription("examly", fields = "Version")),
      error = function(e) "0.4"
    )
  }

  about_content_ui <- function(pkg_ver) {
    tagList(
      tags$h4(T_("Developers", "Geliştiriciler")),
      tags$ul(
        tags$li(HTML(glue::glue(
          'Ahmet Çalışkan - Trakya Üniversitesi - {T_("Contact","İletişim")}: <a href="mailto:ahmetcaliskan1987@gmail.com">ahmetcaliskan1987@gmail.com</a>'
        ))),
        tags$li(HTML(glue::glue(
          'Abdullah Faruk Kılıç - Trakya Üniversitesi - {T_("Contact","İletişim")}: <a href="mailto:afarukkilic@trakya.edu.tr">afarukkilic@trakya.edu.tr</a>'
        )))
      ),
      tags$h4(T_("PkgMeta", "Paket Künyesi")),
      tags$ul(
        tags$li(HTML(glue::glue('{T_("Version","Versiyon")}: {pkg_ver}'))),
        tags$li(HTML(glue::glue('{T_("Github","GitHub")}: <a href="https://github.com/ahmetcaliskan1987/examly" target="_blank">https://github.com/ahmetcaliskan1987/examly</a>')))
      ),
      tags$h4(T_("VersionUpdates", "Yeni Sürümde Güncellenenler")),
      tags$ul(
        tags$li(T_("UpdateModernUI", "Shiny arayüzü modern kart, tablo ve kaydırılabilir panel yapılarıyla yenilendi.")),
        tags$li(T_("UpdateItemTypes", "Madde tipi seçimi daha okunur ve kullanışlı kart düzenine taşındı.")),
        tags$li(T_("UpdateOutcomes", "İsteğe bağlı madde kazanımları eklendi ve en çok yanlış yapılan madde sıralamasında gösterildi.")),
        tags$li(T_("UpdateStats", "Test özetine corrected item-total ayırt edicilik, KR-20 gösterimi, çarpıklık ve basıklık katsayıları eklendi.")),
        tags$li(T_("UpdateReports", "Öğrenci sonuçları, madde sıralaması ve rapor ekranları daha okunur hale getirildi."))
      ),
      tags$h4(T_("Citation", "Akademik Kullanımda Atıf")),
      tags$p("Kılıç, A. F., & Çalışkan, A. (2026, Ocak). Eğitimciler ve Araştırmacılar İçin Açık Kaynak Kodlu Bir Sınav Analizi Aracı: Examly. 2. İstanbul Eğitim Araştırmaları Kongresi.")
    )
  }

  output$about_link <- renderUI({
    dict()
    tags$a(
      id = "about_footer", href = "#",
      HTML(sprintf(
        '<span class="glyphicon glyphicon-info-sign" aria-hidden="true"></span> %s',
        T_("About", "Hakkında")
      ))
    )
  })

  output$reports_about_button <- renderUI({
    dict()
    actionButton("about_reports", T_("About", "Hakkında"), icon = icon("info-sign"))
  })

  show_about_modal <- function() {
    dict()
    pkg_ver <- app_version()
    showModal(modalDialog(
      title     = T_("AppTitle", "examly: Statistical Metrics and Reporting Tool"),
      easyClose = TRUE,
      footer    = modalButton(T_("Close", "Kapat")),
      about_content_ui(pkg_ver),
      tags$p(
        em(HTML(glue::glue('&copy; {format(Sys.Date(), "%Y")} {T_("AllRights","Tüm hakları saklıdır.")}')))
      )
    ))
  }

  observeEvent(input$about_footer, show_about_modal())
  observeEvent(input$about_reports, show_about_modal())

  output$about_info_ui <- renderUI({
    dict()
    tagList(
      h3(T_("About", "Hakkında")),
      about_content_ui(app_version())
    )
  })

  raw_data <- reactive({
    dict()
    req(input$file)

    ext <- tolower(tools::file_ext(input$file$name))

    if (ext %in% c("csv")) {
      delim <- if (input$sep == ";") ";" else ","
      dec   <- if (delim == ";") "," else "."

      try_read <- function(enc) {
        readr::read_delim(
          input$file$datapath,
          delim      = delim,
          col_types  = readr::cols(),
          locale     = readr::locale(encoding = enc, decimal_mark = dec),
          guess_max  = 10000
        )
      }

      df <- tryCatch(
        try_read("UTF-8"),
        error = function(e1) tryCatch(try_read("Windows-1254"), error = function(e2) NULL)
      )

      validate(need(!is.null(df),
                    T_("FileReadError", "Dosya okunamadÃƒâ€Ã‚Â±. CSV biÃƒÆ’Ã‚Â§imini ve ayraÃƒÆ’Ã‚Â§/encoding ayarlarÃƒâ€Ã‚Â±nÃƒâ€Ã‚Â± kontrol edin.")))
      df

    } else if (ext %in% c("xlsx", "xls")) {
      df <- tryCatch(readxl::read_excel(input$file$datapath), error = function(e) NULL)
      validate(need(!is.null(df),
                    T_("FileReadError", "Dosya okunamadÃƒâ€Ã‚Â±. Excel dosyasÃƒâ€Ã‚Â±nda bir sorun olabilir.")))
      df

    } else {
      validate(need(FALSE, T_("UnsupportedFile", "Desteklenmeyen dosya tÃƒÆ’Ã‚Â¼rÃƒÆ’Ã‚Â¼ (sadece CSV/Excel).")))
    }
  })

  output$col_selectors <- renderUI({
    dict()
    df   <- raw_data()
    cols <- names(df)
    default_name_col <- if ("Ad" %in% cols) "Ad" else ""

    tagList(
      selectInput(
        inputId = "name_col",
        label   = T_("NameCol", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci AdÃƒâ€Ã‚Â± SÃƒÆ’Ã‚Â¼tunu (isteÃƒâ€Ã…Â¸e baÃƒâ€Ã…Â¸lÃƒâ€Ã‚Â±)"),
        choices = setNames(
          c("", cols),
          c(paste0("(", T_("None", "Yok"), ")"), cols)
        ),
        selected = default_name_col
      ),
      selectizeInput(
        inputId = "item_cols",
        label   = T_("ItemCols", "Madde SÃƒÆ’Ã‚Â¼tunlarÃƒâ€Ã‚Â± (maddeler)"),
        choices = setdiff(cols, default_name_col),
        multiple = TRUE,
        selected = setdiff(cols, c("Ad","Ãƒâ€Ã‚Â°sim","Name"))
      )
    )
  })

  observeEvent(input$name_col, {
    df <- raw_data()
    cols <- names(df)
    name_col <- examly::`%||%`(input$name_col, "")
    item_choices <- setdiff(cols, name_col)
    item_selected <- intersect(
      setdiff(examly::norm_cols(input$item_cols), name_col),
      item_choices
    )

    updateSelectizeInput(
      session,
      "item_cols",
      choices = item_choices,
      selected = item_selected
    )
  }, ignoreInit = FALSE)

  data_items <- reactive({
    df <- raw_data(); req(examly::norm_cols(input$item_cols))
    sel <- setdiff(examly::norm_cols(input$item_cols), input$name_col)
    sel <- intersect(sel, names(df))
    if(length(sel) == 0) return(df[, 0, drop = FALSE])
    as.data.frame(df[, sel, drop=FALSE])
  })

  outcome_input_id <- function(index) {
    paste0("lo_", index)
  }

  output$learning_outcomes_ui <- renderUI({
    dict()
    cols <- examly::norm_cols(input$item_cols)
    if (length(cols) == 0) return(NULL)
    header_item <- T_("ItemCol", "Madde")
    header_outcome <- T_("LearningOutcome", "KazanÃƒâ€Ã‚Â±m")
    rows <- lapply(seq_along(cols), function(i) {
      nm <- cols[[i]]
      tags$tr(
        tags$td(nm),
        tags$td(textInput(outcome_input_id(i), label = NULL, value = "", width = "100%",
                          placeholder = T_("LearningOutcomePlaceholder", "Ãƒâ€Ã‚Â°steÃƒâ€Ã…Â¸e baÃƒâ€Ã…Â¸lÃƒâ€Ã‚Â± kazanÃƒâ€Ã‚Â±m yazÃƒâ€Ã‚Â±n")))
      )
    })
    tagList(
      h4(T_("LearningOutcomesHeader", "Madde KazanÃƒâ€Ã‚Â±mlarÃƒâ€Ã‚Â±")),
      tags$p(class = "muted", T_("LearningOutcomesNote", "Ã„Â°steÃ„Å¸e baÃ„Å¸lÃ„Â±dÃ„Â±r. Girilirse en ÃƒÂ§ok yanlÃ„Â±Ã…Å¸ yapÃ„Â±lan madde sÃ„Â±ralamasÃ„Â±nda ilgili kazanÃ„Â±m da gÃƒÂ¶sterilir.")),
      tags$div(
        class = "outcome-grid-wrap",
        tags$table(
          class = "outcome-grid",
          tags$thead(tags$tr(tags$th(header_item), tags$th(header_outcome))),
          tags$tbody(rows)
        )
      )
    )
  })

  item_outcomes <- reactive({
    cols <- examly::norm_cols(input$item_cols)
    vals <- vapply(seq_along(cols), function(i) {
      stringr::str_squish(as.character(examly::`%||%`(input[[outcome_input_id(i)]], "")))
    }, character(1))
    names(vals) <- cols
    vals
  })

  student_names <- reactive({
    dict()
    df <- raw_data()
    n <- nrow(df)

    base <- T_("StudentWord", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci")

    if (!isTruthy(input$name_col) || input$name_col == "") {
      nm <- sprintf("%s_%03d", base, seq_len(n))
    } else {
      nm <- as.character(df[[input$name_col]])
      nm <- stringr::str_squish(nm)
      miss <- is.na(nm) | nm == ""
      if (any(miss)) {
        nm[miss] <- sprintf("%s_%03d", base, which(miss))
      }
    }
    nm <- make.unique(nm, sep = "_")
    nm
  })

  output$dimensions_info <- renderTable({
    dict()
    df <- data_items()
    tib <- tibble::tibble(
      StudentsCount = nrow(df),
      ItemsCount    = ncol(df)
    )
    names(tib) <- c(
      T_("StudentsCount", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci (katÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â±mcÃƒâ€Ã‚Â±) sayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"),
      T_("ItemsCount",    "Soru (madde) sayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
    )
    tib
  }, align = "c")

  output$sample_head <- renderTable({
    dict()
    nm_col <- T_("NameHeader", "Ad")
    out <- cbind(setNames(data.frame(student_names()), nm_col), data_items())
    utils::head(out, 5)
  }, align = "c")

  mc_items_use <- reactive({
    taken <- union(
      union(examly::`%||%`(input$tf_items, character(0)),
            examly::`%||%`(input$lc_items, character(0))),
      examly::`%||%`(input$bin_items, character(0))
    )
    base_cols <- examly::`%||%`(examly::norm_cols(input$item_cols), character(0))
    if(!is.null(input$mc_items) && length(input$mc_items)>0) input$mc_items else setdiff(base_cols, taken)
  })
  tf_items_use  <- reactive({ examly::`%||%`(input$tf_items,  character(0)) })
  lc_items_use  <- reactive({ examly::`%||%`(input$lc_items,  character(0)) })
  bin_items_use <- reactive({ examly::`%||%`(input$bin_items, character(0)) })

  # --- SELECT ALL BUTTONS (RENDERED IN SERVER) ---
  create_header_with_link <- function(label_text, link_id, link_text) {
    tags$div(
      class = "item-type-card-head",
      tags$div(class = "item-type-card-title", label_text),
      actionLink(link_id, link_text, class = "select-all-link")
    )
  }

  output$mc_item_selector <- renderUI({
    req(input$item_cols); dict()
    sel_now <- isolate(input$mc_items)
    tags$div(
      class = "item-type-card",
      create_header_with_link(T_("MCItems","ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli Maddeler"), "sel_all_mc", T_("SelectAll", "TÃƒÆ’Ã‚Â¼mÃƒÆ’Ã‚Â¼nÃƒÆ’Ã‚Â¼ SeÃƒÆ’Ã‚Â§")),
      checkboxGroupInput("mc_items", label = NULL,
                         choices=input$item_cols, selected=sel_now %||% mc_items_use())
    )
  })

  output$tf_item_selector <- renderUI({
    req(input$item_cols); dict()
    sel_now <- isolate(input$tf_items)
    tags$div(
      class = "item-type-card",
      create_header_with_link(T_("TFItems","DoÃƒâ€Ã…Â¸ru/YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸ Maddeler"), "sel_all_tf", T_("SelectAll", "TÃƒÆ’Ã‚Â¼mÃƒÆ’Ã‚Â¼nÃƒÆ’Ã‚Â¼ SeÃƒÆ’Ã‚Â§")),
      checkboxGroupInput("tf_items", label = NULL,
                         choices=input$item_cols, selected=sel_now %||% tf_items_use())
    )
  })

  output$bin_item_selector <- renderUI({
    req(input$item_cols); dict()
    sel_now <- isolate(input$bin_items)
    tags$div(
      class = "item-type-card",
      create_header_with_link(T_("BINItems","1-0 Kodlu Maddeler"), "sel_all_bin", T_("SelectAll", "TÃƒÆ’Ã‚Â¼mÃƒÆ’Ã‚Â¼nÃƒÆ’Ã‚Â¼ SeÃƒÆ’Ã‚Â§")),
      checkboxGroupInput("bin_items", label = NULL,
                         choices=input$item_cols, selected=sel_now %||% bin_items_use())
    )
  })

  output$lc_item_selector <- renderUI({
    req(input$item_cols); dict()
    sel_now <- isolate(input$lc_items)
    tags$div(
      class = "item-type-card",
      create_header_with_link(T_("LCItems","Uzun CevaplÃƒâ€Ã‚Â± Maddeler"), "sel_all_lc", T_("SelectAll", "TÃƒÆ’Ã‚Â¼mÃƒÆ’Ã‚Â¼nÃƒÆ’Ã‚Â¼ SeÃƒÆ’Ã‚Â§")),
      checkboxGroupInput("lc_items", label = NULL,
                         choices=input$item_cols, selected=sel_now %||% lc_items_use())
    )
  })

  # --- OBSERVERS FOR SELECT ALL ---
  observeEvent(input$sel_all_mc, {
    req(input$item_cols)
    updateCheckboxGroupInput(session, "mc_items", choices = input$item_cols, selected = input$item_cols)
  })
  observeEvent(input$sel_all_tf, {
    req(input$item_cols)
    updateCheckboxGroupInput(session, "tf_items", choices = input$item_cols, selected = input$item_cols)
  })
  observeEvent(input$sel_all_bin, {
    req(input$item_cols)
    updateCheckboxGroupInput(session, "bin_items", choices = input$item_cols, selected = input$item_cols)
  })
  observeEvent(input$sel_all_lc, {
    req(input$item_cols)
    updateCheckboxGroupInput(session, "lc_items", choices = input$item_cols, selected = input$item_cols)
  })
  # ----------------------------------------

  mc_levels <- reactive({
    ch <- as.integer(input$mc_choice_count)[1]; if(is.na(ch)) ch <- 5
    if(ch==5) c("A","B","C","D","E") else if(ch==4) c("A","B","C","D") else c("A","B","C")
  })

  output$mc_answer_table <- renderUI({
    dict();
    hdrs <- mc_items_use()
    if(length(hdrs)==0){
      return(tags$div(class="muted", T_("ThisSectionEmpty","(Bu bÃƒÆ’Ã‚Â¶lÃƒÆ’Ã‚Â¼mde madde yok)")))
    }
    n_txt <- as.character(length(hdrs))
    help1 <- gsub("__N__", n_txt, T_("MCKeyPasteHelp1",
                                     "AÃƒâ€¦Ã…Â¸aÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±daki __N__ adet ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli madde iÃƒÆ’Ã‚Â§in cevaplarÃƒâ€Ã‚Â± SIRA GÃƒÆ’Ã¢â‚¬â€œZETEREK yapÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸tÃƒâ€Ã‚Â±rÃƒâ€Ã‚Â±n."))
    help2 <- T_("MCKeyPasteHelp2","GiriÃƒâ€¦Ã…Â¸ SÃƒâ€Ã‚Â±rasÃƒâ€Ã‚Â±:")

    tagList(
      helpText(help1),
      helpText(strong(paste(help2, paste(hdrs, collapse=", ")))),
      textAreaInput(
        "mc_key_paste",
        T_("MCKeyHeader","ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli AnahtarÃƒâ€Ã‚Â± (SÃƒâ€Ã‚Â±rayla YapÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸tÃƒâ€Ã‚Â±rÃƒâ€Ã‚Â±n)"),
        rows = 3,
        placeholder = T_("MCKeyPlaceholder","ÃƒÆ’Ã¢â‚¬â€œrn: A,B,C,D,E,A,B...")
      )
    )
  })

  output$tf_answer_table <- renderUI({
    dict()
    hdrs <- tf_items_use()
    if(length(hdrs)==0){
      return(tags$div(class="muted", T_("ThisSectionNoTF","(DY olarak iÃƒâ€¦Ã…Â¸aretlenen madde yok)")))
    }
    n_txt <- as.character(length(hdrs))
    help1 <- gsub("__N__", n_txt, T_("TFKeyPasteHelp1",
                                     "AÃƒâ€¦Ã…Â¸aÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±daki __N__ adet D/Y madde iÃƒÆ’Ã‚Â§in cevaplarÃƒâ€Ã‚Â± SIRA GÃƒÆ’Ã¢â‚¬â€œZETEREK yapÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸tÃƒâ€Ã‚Â±rÃƒâ€Ã‚Â±n."))
    help2 <- T_("MCKeyPasteHelp2","GiriÃƒâ€¦Ã…Â¸ SÃƒâ€Ã‚Â±rasÃƒâ€Ã‚Â±:")
    tagList(
      helpText(help1),
      helpText(strong(paste(help2, paste(hdrs, collapse=", ")))),
      textAreaInput(
        "tf_key_paste",
        T_("TFKeyHeader","DoÃƒâ€Ã…Â¸ru/YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸ AnahtarÃƒâ€Ã‚Â± (SÃƒâ€Ã‚Â±rayla YapÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸tÃƒâ€Ã‚Â±rÃƒâ€Ã‚Â±n)"),
        rows = 2,
        placeholder = T_("TFKeyPlaceholder","ÃƒÆ’Ã¢â‚¬â€œrn: D,Y,D,D...")
      )
    )
  })

  output$ui_distribute_weights_btn <- renderUI({
    dict()
    actionButton("distribute_weights_btn",
                 T_("DistributeWeights", "PuanlarÃƒâ€Ã‚Â± 100 ÃƒÆ’Ã…â€œzerinden EÃƒâ€¦Ã…Â¸it DaÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±t"),
                 icon = icon("balance-scale"),
                 style = "font-size:12px; padding:4px 8px; margin-bottom:10px;")
  })

  observeEvent(input$distribute_weights_btn, {
    req(examly::norm_cols(input$item_cols))
    cols <- examly::norm_cols(input$item_cols)
    n <- length(cols)
    if(n > 0){
      val <- 100 / n
      for(nm in cols){
        updateNumericInput(session, paste0("w_", nm), value = val)
      }
    }
  })

  output$item_weight_table <- renderUI({
    req(examly::norm_cols(input$item_cols))
    hdrs <- examly::norm_cols(input$item_cols)
    inputs <- lapply(hdrs, function(nm){
      tagAppendAttributes(numericInput(paste0("w_", nm), NULL, value = 1, min = 0, step = 0.5, width=NULL),
                          class="grid-input-num")
    })
    tags$table(class="compact-grid weights-grid",
               tags$thead(tags$tr(lapply(hdrs, tags$th))),
               tags$tbody(tags$tr(lapply(inputs, function(x) tags$td(x))))
    )
  })

  observeEvent(input$mc_items, {
    mc <- examly::`%||%`(input$mc_items,  character(0))
    tf <- examly::`%||%`(input$tf_items,  character(0))
    lc <- examly::`%||%`(input$lc_items,  character(0))
    bn <- examly::`%||%`(input$bin_items, character(0))
    new_tf <- setdiff(tf, mc); new_lc <- setdiff(lc, mc); new_bn <- setdiff(bn, mc)
    if(!identical(new_tf, tf)) updateCheckboxGroupInput(session, "tf_items", selected = new_tf)
    if(!identical(new_lc, lc)) updateCheckboxGroupInput(session, "lc_items", selected = new_lc)
    if(!identical(new_bn, bn)) updateCheckboxGroupInput(session, "bin_items", selected = new_bn)
  }, ignoreInit = TRUE)
  observeEvent(input$tf_items, {
    mc <- examly::`%||%`(input$mc_items,  character(0))
    tf <- examly::`%||%`(input$tf_items,  character(0))
    lc <- examly::`%||%`(input$lc_items,  character(0))
    bn <- examly::`%||%`(input$bin_items, character(0))
    new_mc <- setdiff(mc, tf); new_lc <- setdiff(lc, tf); new_bn <- setdiff(bn, tf)
    if(!identical(new_mc, mc)) updateCheckboxGroupInput(session, "mc_items", selected = new_mc)
    if(!identical(new_lc, lc)) updateCheckboxGroupInput(session, "lc_items", selected = new_lc)
    if(!identical(new_bn, bn)) updateCheckboxGroupInput(session, "bin_items", selected = new_bn)
  }, ignoreInit = TRUE)
  observeEvent(input$bin_items, {
    mc <- examly::`%||%`(input$mc_items,  character(0))
    tf <- examly::`%||%`(input$tf_items,  character(0))
    lc <- examly::`%||%`(input$lc_items,  character(0))
    bn <- examly::`%||%`(input$bin_items, character(0))
    new_mc <- setdiff(mc, bn); new_tf <- setdiff(tf, bn); new_lc <- setdiff(lc, bn)
    if(!identical(new_mc, mc)) updateCheckboxGroupInput(session, "mc_items", selected = new_mc)
    if(!identical(new_tf, tf)) updateCheckboxGroupInput(session, "tf_items", selected = new_tf)
    if(!identical(new_lc, lc)) updateCheckboxGroupInput(session, "lc_items", selected = new_lc)
  }, ignoreInit = TRUE)
  observeEvent(input$lc_items, {
    mc <- examly::`%||%`(input$mc_items,  character(0))
    tf <- examly::`%||%`(input$tf_items,  character(0))
    lc <- examly::`%||%`(input$lc_items,  character(0))
    bn <- examly::`%||%`(input$bin_items, character(0))
    new_mc <- setdiff(mc, lc); new_tf <- setdiff(tf, lc); new_bn <- setdiff(bn, lc)
    if(!identical(new_mc, mc)) updateCheckboxGroupInput(session, "mc_items", selected = new_mc)
    if(!identical(new_tf, tf)) updateCheckboxGroupInput(session, "tf_items", selected = new_tf)
    if(!identical(new_bn, bn)) updateCheckboxGroupInput(session, "bin_items", selected = new_bn)
  }, ignoreInit = TRUE)

  answer_keys <- reactive({
    mc_names <- mc_items_use()
    tf_names <- tf_items_use()
    lv <- mc_levels()

    # --- MC (ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli) BÃƒÆ’Ã‚Â¶lÃƒÆ’Ã‚Â¼mÃƒÆ’Ã‚Â¼ ---
    mc <- setNames(rep(NA_character_, length(mc_names)), mc_names)
    if (length(mc_names) > 0 && isTruthy(input$mc_key_paste)) {
      raw_text_mc <- input$mc_key_paste

      # DEÃƒâ€Ã‚ÂÃƒâ€Ã‚Â°Ãƒâ€¦Ã‚ÂÃƒâ€Ã‚Â°KLÃƒâ€Ã‚Â°K BURADA:
      # EÃƒâ€Ã…Â¸er metinde virgÃƒÆ’Ã‚Â¼l veya noktalÃƒâ€Ã‚Â± virgÃƒÆ’Ã‚Â¼l varsa ayraÃƒÆ’Ã‚Â§ olarak onlarÃƒâ€Ã‚Â± kullan
      if (grepl("[,;]", raw_text_mc)) {
        keys_pasted <- unlist(strsplit(raw_text_mc, "[,;\\n\\t]+"))
      } else {
        # Yoksa (ÃƒÆ’Ã‚Â¶rneÃƒâ€Ã…Â¸in A B C D veya ABCD), tÃƒÆ’Ã‚Â¼m boÃƒâ€¦Ã…Â¸luklarÃƒâ€Ã‚Â± (space, tab, enter) sil
        # ve kalan harfleri tek tek ayÃƒâ€Ã‚Â±r.
        clean_text <- gsub("\\s+", "", raw_text_mc)
        keys_pasted <- unlist(strsplit(clean_text, ""))
      }

      keys_pasted <- trimws(keys_pasted)
      keys_pasted <- keys_pasted[keys_pasted != ""]
      keys_pasted <- toupper(keys_pasted)
      n_keys <- length(keys_pasted)
      n_items <- length(mc_names)
      n_match <- min(n_keys, n_items)
      if (n_match > 0) {
        valid_keys <- keys_pasted[1:n_match]
        valid_keys[!valid_keys %in% lv] <- NA_character_
        mc[1:n_match] <- valid_keys
      }
    }

    # --- TF (DoÃƒâ€Ã…Â¸ru/YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸) BÃƒÆ’Ã‚Â¶lÃƒÆ’Ã‚Â¼mÃƒÆ’Ã‚Â¼ ---
    tf <- setNames(rep(NA_character_, length(tf_names)), tf_names)
    if (length(tf_names) > 0 && isTruthy(input$tf_key_paste)) {
      raw_text_tf <- input$tf_key_paste

      # DEÃƒâ€Ã‚ÂÃƒâ€Ã‚Â°Ãƒâ€¦Ã‚ÂÃƒâ€Ã‚Â°KLÃƒâ€Ã‚Â°K BURADA (AynÃƒâ€Ã‚Â± mantÃƒâ€Ã‚Â±k):
      if (grepl("[,;]", raw_text_tf)) {
        keys_pasted_tf <- unlist(strsplit(raw_text_tf, "[,;\\n\\t]+"))
      } else {
        clean_text_tf <- gsub("\\s+", "", raw_text_tf)
        keys_pasted_tf <- unlist(strsplit(clean_text_tf, ""))
      }

      keys_pasted_tf <- trimws(keys_pasted_tf)
      keys_pasted_tf <- keys_pasted_tf[keys_pasted_tf != ""]
      keys_pasted_tf <- toupper(keys_pasted_tf)
      n_keys_tf <- length(keys_pasted_tf)
      n_items_tf <- length(tf_names)
      n_match_tf <- min(n_keys_tf, n_items_tf)
      if (n_match_tf > 0) {
        valid_keys_tf <- keys_pasted_tf[1:n_match_tf]
        valid_keys_tf[!valid_keys_tf %in% c("D", "Y")] <- NA_character_
        tf[1:n_match_tf] <- valid_keys_tf
      }
    }
    list(mc = mc, tf = tf)
  })

  item_weights <- reactive({
    cols <- examly::`%||%`(examly::norm_cols(input$item_cols), character(0))
    w <- setNames(rep(1, length(cols)), cols)
    for(nm in cols){
      v <- suppressWarnings(as.numeric(input[[paste0("w_", nm)]]))
      if(length(v) == 0 || is.na(v)) next
      if(v >= 0) w[[nm]] <- v
    }
    w
  })

  scored_items <- reactive({
    df <- data_items(); keys <- answer_keys()
    if (is.null(df) || (length(examly::`%||%`(examly::norm_cols(input$item_cols), character(0))) == 0)) return(NULL)

    all_names <- examly::norm_cols(input$item_cols)
    sc_bin <- matrix(NA_real_, nrow=nrow(df), ncol=length(all_names), dimnames=list(NULL, all_names)) %>% as.data.frame()
    lc_raw <- matrix(NA_real_, nrow=nrow(df), ncol=length(all_names), dimnames=list(NULL, all_names)) %>% as.data.frame()

    for(nm in lc_items_use()){
      if(nm %in% names(df)){
        lc_raw[[nm]] <- examly::parse_lc_raw(df[[nm]])
        sc_bin[[nm]] <- NA_real_
      }
    }

    for(nm in bin_items_use()){
      if(nm %in% names(df)){
        sc_bin[[nm]] <- examly::parse_lc_bin(df[[nm]])
      }
    }

    for(nm in setdiff(all_names, union(union(mc_items_use(), tf_items_use()),
                                       union(lc_items_use(), bin_items_use())))){
      col <- df[[nm]]
      if(examly::is_scored_01(col)){
        sc_bin[[nm]] <- examly::parse_lc_bin(col)
      }
    }

    for(nm in names(keys$mc)){
      if(isTruthy(keys$mc[[nm]]) && nm %in% names(df)){
        sc_bin[[nm]] <- examly::parse_mc_bin(df[[nm]], keys$mc[[nm]])
      }
    }

    for(nm in names(keys$tf)){
      if(isTruthy(keys$tf[[nm]]) && nm %in% names(df)){
        sc_bin[[nm]] <- examly::parse_tf_bin(df[[nm]], keys$tf[[nm]])
      }
    }

    list(sc_bin = sc_bin, lc_raw = lc_raw)
  })

  output$max_score_banner <- renderUI({
    s <- test_summary()
    if (is.null(s) || is.null(s$max_weighted)) return(NULL)
    tags$div(class = "banner",
             HTML(paste0("<b>", T_("MaxScore","AlÃƒâ€Ã‚Â±nabilecek Maksimum Puan:"), "</b> ",
                         round(s$max_weighted, 3)))
    )
  })

  itemexam_quant_reactive <- reactive({ examly::get_itemexam_quant() })

  test_summary <- reactive({
    dict()
    si <- scored_items()
    validate(need(!is.null(si), T_("NeedItemTypesKeys", "LÃƒÆ’Ã‚Â¼tfen madde tÃƒÆ’Ã‚Â¼rlerini ve anahtarlarÃƒâ€Ã‚Â± girin.")))

    sc_bin <- si$sc_bin
    lc_raw <- si$lc_raw
    n_stu <- nrow(sc_bin)

    w <- item_weights()
    all_items <- colnames(sc_bin)
    w_use <- examly::`%||%`(w[all_items], rep(1, length(all_items)))
    names(w_use) <- all_items

    Y <- as.data.frame(matrix(NA_real_, nrow = n_stu, ncol = length(all_items), dimnames = list(NULL, all_items)))

    mc_tf_bin_names <- union(union(mc_items_use(), tf_items_use()), bin_items_use())
    lc_names    <- lc_items_use()
    other_names <- setdiff(all_items, union(mc_tf_bin_names, lc_names))

    for(nm in intersect(mc_tf_bin_names, all_items)){
      Y[[nm]] <- sc_bin[[nm]] * w_use[[nm]]
    }

    for(nm in intersect(lc_names, all_items)){
      v <- lc_raw[[nm]]; mx <- as.numeric(w_use[[nm]])
      if(!is.na(mx)) v <- pmin(pmax(v, 0), mx)
      Y[[nm]] <- v
    }

    for(nm in other_names){
      Y[[nm]] <- sc_bin[[nm]] * w_use[[nm]]
    }

    weighted_total <- rowSums(Y, na.rm = TRUE)
    max_weighted <- sum(w_use[intersect(union(mc_tf_bin_names, lc_names), names(w_use))], na.rm = TRUE) +
      sum(w_use[intersect(setdiff(all_items, union(mc_tf_bin_names, lc_names)), names(w_use))], na.rm = TRUE)

    p_all <- rep(NA_real_, length(all_items)); names(p_all) <- all_items
    for(nm in all_items){
      mj <- w_use[[nm]]
      if(!is.na(mj) && mj>0){
        p_all[[nm]] <- mean(Y[[nm]], na.rm = TRUE) / mj
      } else p_all[[nm]] <- NA_real_
    }

    corrected_item_total <- sapply(seq_len(ncol(Y)), function(j){
      yj <- Y[[j]]
      if(all(is.na(yj))) return(NA_real_)
      examly::pbiserial_rest(yj, weighted_total - ifelse(is.na(yj), 0, yj))
    })
    names(corrected_item_total) <- all_items
    rjx <- corrected_item_total

    avg_p_psych <- NA_real_
    avg_r_psych <- NA_real_
    ul27_disc <- rep(NA_real_, length(all_items)); names(ul27_disc) <- all_items

    if(n_stu >= 4){
      k_ul <- ceiling(n_stu * 0.27)
      ord <- order(weighted_total)
      lower_idx <- ord[1:k_ul]
      upper_idx <- ord[(n_stu-k_ul+1):n_stu]

      for(nm in all_items){
        mj <- w_use[[nm]]
        col_vals <- Y[[nm]]
        if(!is.na(mj) && mj > 0 && !all(is.na(col_vals))){
          mean_upper <- mean(col_vals[upper_idx], na.rm=TRUE)
          mean_lower <- mean(col_vals[lower_idx], na.rm=TRUE)
          p_upper <- mean_upper / mj
          p_lower <- mean_lower / mj
          ul27_disc[[nm]] <- p_upper - p_lower
        }
      }
    }

    only_dicho_names <- intersect(colnames(sc_bin), union(mc_tf_bin_names, other_names))
    if(requireNamespace("psychometric", quietly=TRUE) && length(only_dicho_names) > 0){
      dicho_mat <- sc_bin[, only_dicho_names, drop=FALSE]
      try({
        res_psych <- psychometric::item.exam(dicho_mat, discrim = TRUE)
        avg_p_psych <- mean(res_psych$Difficulty, na.rm=TRUE)
        avg_r_psych <- mean(res_psych$Discrimination, na.rm=TRUE)
      }, silent=TRUE)
    }

    if(is.na(avg_p_psych)) avg_p_psych <- mean(p_all, na.rm=TRUE)
    avg_r_psych <- mean(rjx, na.rm=TRUE)

    avg_ul27 <- mean(ul27_disc, na.rm = TRUE)

    valid_cols <- which(colSums(!is.na(Y)) > 1)
    k_eff <- length(valid_cols)
    col_vars <- if(k_eff>0) sapply(valid_cols, function(j) stats::var(Y[[j]], na.rm = TRUE)) else numeric(0)
    var_total <- stats::var(weighted_total, na.rm = TRUE)
    total_centered <- weighted_total - mean(weighted_total, na.rm = TRUE)
    total_sd <- stats::sd(weighted_total, na.rm = TRUE)
    total_skewness <- if(!is.na(total_sd) && total_sd > 0) mean(total_centered^3, na.rm = TRUE) / (total_sd^3) else NA_real_
    total_kurtosis <- if(!is.na(total_sd) && total_sd > 0) mean(total_centered^4, na.rm = TRUE) / (total_sd^4) - 3 else NA_real_

    # Cronbach Alfa HesabÃƒâ€Ã‚Â±
    alpha <- if (k_eff > 1 && !is.na(var_total) && var_total > 0) (k_eff/(k_eff - 1)) * (1 - sum(col_vars, na.rm = TRUE)/var_total) else NA_real_

    is_all_binary <- length(lc_names) == 0 &&
      all(apply(sc_bin, 2, function(x) all(x %in% c(0,1,NA))), na.rm=TRUE)
    is_all_mc <- length(all_items) > 0 &&
      setequal(all_items, mc_items_use()) &&
      length(tf_items_use()) == 0 &&
      length(bin_items_use()) == 0 &&
      length(lc_items_use()) == 0

    # DÃƒÆ’Ã…â€œZELTME: KR-20 hesabÃƒâ€Ã‚Â±
    # EÃƒâ€Ã…Â¸er veri binary ise KR-20 matematiksel olarak Alfa ile aynÃƒâ€Ã‚Â±dÃƒâ€Ã‚Â±r.
    # Manuel hesaplamadaki olasÃƒâ€Ã‚Â± 'pq toplamÃƒâ€Ã‚Â± 0' hatasÃƒâ€Ã‚Â±nÃƒâ€Ã‚Â± ÃƒÆ’Ã‚Â¶nlemek iÃƒÆ’Ã‚Â§in doÃƒâ€Ã…Â¸rudan Alfa deÃƒâ€Ã…Â¸erini kullanÃƒâ€Ã‚Â±yoruz.
    kr20_val <- NA_real_
    if(is_all_binary && k_eff > 1){
      kr20_val <- alpha
    }

    list(
      n_stu = n_stu,
      median = stats::median(weighted_total, na.rm=TRUE),
      mode   = suppressWarnings(examly::d_mode(weighted_total)),
      mean   = mean(weighted_total, na.rm=TRUE),
      sd     = stats::sd(weighted_total, na.rm=TRUE),
      skewness = total_skewness,
      kurtosis = total_kurtosis,
      avg_p  = mean(p_all, na.rm=TRUE),
      avg_r  = mean(rjx, na.rm=TRUE),
      avg_ul27 = avg_ul27,
      avg_p_psych = avg_p_psych,
      avg_r_psych = avg_r_psych,
      alpha  = alpha,
      kr20   = kr20_val,
      is_all_binary = is_all_binary,
      is_all_mc = is_all_mc,
      max_weighted = max_weighted,

      total_unweighted = rowSums(sc_bin, na.rm = TRUE),
      p = p_all, rjx = rjx, ul27 = ul27_disc,
      sc_bin = sc_bin,
      lc_raw = lc_raw,
      Y = Y,
      weighted_total = weighted_total,
      max_weighted = max_weighted,
      weights = w_use
    )
  })

  student_scores_processed <- reactive({
    dict()
    s <- test_summary()
    validate(need(!is.null(s) && !is.null(s$weighted_total) && !is.null(s$max_weighted) && s$max_weighted > 0,
                  T_("WaitingForScores", "Puanlar hesaplanÃƒâ€Ã‚Â±yor veya maksimum puan 0...")))

    percent_scores <- (s$weighted_total / s$max_weighted) * 100
    count_pass_50 <- sum(percent_scores >= 50, na.rm = TRUE)
    score_labels_i18n <- c(
      T_("Grade.Fail", "0-49.99 GeÃƒÆ’Ã‚Â§mez"),
      T_("Grade.Pass", "50-59.99 GeÃƒÆ’Ã‚Â§er"),
      T_("Grade.Medium", "60-69.99 Orta"),
      T_("Grade.Good", "70-84.99 Ãƒâ€Ã‚Â°yi"),
      T_("Grade.Excellent", "85-100 Pekiyi")
    )
    score_breaks <- c(-Inf, 49.99, 59.99, 69.99, 84.99, 100)
    grades <- cut(percent_scores, breaks = score_breaks, labels = score_labels_i18n, right = TRUE)
    df_grades <- data.frame(Grades = grades)
    list(
      count_pass_50 = count_pass_50,
      n_total = length(percent_scores),
      df_grades = df_grades,
      grades_i18n = score_labels_i18n
    )
  })

  item_highlights <- reactive({
    input$lang; dict()
    s <- test_summary()
    validate(need(!is.null(s) && !is.null(s$sc_bin), T_("WaitingForScores", "Puanlar hesaplanÃƒâ€Ã‚Â±yor...")))
    sc_bin <- s$sc_bin
    n_total <- s$n_stu
    binary_items <- setdiff(colnames(sc_bin), lc_items_use())
    if (length(binary_items) == 0 || n_total == 0) {
      return(list(most_correct = NA_character_, correct_count = 0,
                  most_wrong = NA_character_,   wrong_count = 0,
                  most_blank = NA_character_,   blank_count = 0,
                  n_total = n_total))
    }
    sc_bin_binary_only <- sc_bin[, binary_items, drop = FALSE]
    n_correct <- colSums(sc_bin_binary_only == 1, na.rm = TRUE)
    n_wrong   <- colSums(sc_bin_binary_only == 0, na.rm = TRUE)
    n_blank   <- colSums(is.na(sc_bin_binary_only), na.rm = TRUE)
    item_most_correct <- if(all(is.na(n_correct)) || length(n_correct)==0 || all(n_correct == 0)) NA_character_ else names(n_correct)[which.max(n_correct)]
    item_most_wrong   <- if(all(is.na(n_wrong)) || length(n_wrong)==0 || all(n_wrong == 0)) NA_character_ else names(n_wrong)[which.max(n_wrong)]
    item_most_blank   <- if(all(is.na(n_blank)) || length(n_blank)==0 || all(n_blank == 0)) NA_character_ else names(n_blank)[which.max(n_blank)]
    list(
      most_correct = item_most_correct,
      correct_count = if(is.na(item_most_correct)) 0 else n_correct[item_most_correct],
      most_wrong = item_most_wrong,
      wrong_count = if(is.na(item_most_wrong)) 0 else n_wrong[item_most_wrong],
      most_blank = item_most_blank,
      blank_count = if(is.na(item_most_blank)) 0 else n_blank[item_most_blank],
      n_total = n_total
    )
  })

  output$weighted_mean_text <- renderText({
    s <- test_summary()
    paste0(T_("WeightedMean", "SÃƒâ€Ã‚Â±nav OrtalamasÃƒâ€Ã‚Â±:"), " ", round(mean(s$weighted_total, na.rm = TRUE), 3))
  })

  output$overall_comment <- renderUI({
    dict()
    s <- test_summary()
    keyP <- if (is.na(s$avg_p)) NULL else if (s$avg_p < .4) "Overall_ptxt_low"
    else if (s$avg_p > .8) "Overall_ptxt_high" else "Overall_ptxt_mid"
    keyR <- if (is.na(s$avg_r)) NULL else if (s$avg_r < .2) "Overall_rtxt_low"
    else if (s$avg_r > .4) "Overall_rtxt_high" else "Overall_rtxt_mid"
    ptxt <- if (is.null(keyP)) "" else T_(keyP, "")
    rtxt <- if (is.null(keyR)) "" else T_(keyR, "")
    p_state <- if (is.na(s$avg_p)) "NA" else if (s$avg_p < .4) T_("Difficult", "Zor")
    else if (s$avg_p > .8) T_("Easy", "Kolay") else T_("Medium", "Orta")
    r_state <- if (is.na(s$avg_r)) "NA" else if (s$avg_r < .2) T_("Low", "Dusuk")
    else if (s$avg_r > .4) T_("High", "Yuksek") else T_("Medium", "Orta")
    p_class <- if (!is.na(s$avg_p) && s$avg_p >= .4 && s$avg_p <= .8) "value-pill-good" else "value-pill-risk"
    r_class <- if (!is.na(s$avg_r) && s$avg_r >= .3) "value-pill-good" else "value-pill-risk"

    tags$table(
      class = "table table-condensed comment-table",
      tags$tbody(
        tags$tr(
          tags$td(T_("AvgDifficulty", "Ortalama Gucluk (p)")),
          tags$td(class = "comment-state", tags$span(class = paste("value-pill", p_class), p_state)),
          tags$td(ptxt)
        ),
        tags$tr(
          tags$td(T_("AvgDiscrimination", "Ortalama Madde Ayirt Ediciligi")),
          tags$td(class = "comment-state", tags$span(class = paste("value-pill", r_class), r_state)),
          tags$td(rtxt)
        )
      )
    )
  })

  output$summary_cards <- renderUI({
    dict()
    s <- test_summary()
    rel_val <- if((isTRUE(s$is_all_mc) || isTRUE(s$is_all_binary)) && !is.na(s$kr20)) s$kr20 else s$alpha
    rel_lbl <- if((isTRUE(s$is_all_mc) || isTRUE(s$is_all_binary)) && !is.na(s$kr20)) T_("KR20", "KR-20") else T_("CronbachAlpha", "Cronbach Alfa")
    card <- function(label, value, note = NULL) {
      tags$div(
        class = "metric-card",
        tags$div(class = "metric-label", label),
        tags$div(class = "metric-value", value),
        if (!is.null(note)) tags$div(class = "metric-note", note)
      )
    }
    tags$div(
      class = "metric-grid",
      card(T_("StudentsCount", "Ogrenci Sayisi"), s$n_stu),
      card(T_("WeightedExamMean", "Sinav Ortalamasi"), round(mean(s$weighted_total, na.rm = TRUE), 3),
           paste(T_("MaxWeighted", "Alinabilecek Maksimum Puan"), round(s$max_weighted, 3))),
      card(T_("AvgDifficulty", "Ortalama Gucluk (p)"), round(s$avg_p, 3)),
      card(rel_lbl, round(rel_val, 3))
    )
  })

  output$test_summary_table <- renderUI({
    s <- test_summary()

    rel_val <- if((isTRUE(s$is_all_mc) || isTRUE(s$is_all_binary)) && !is.na(s$kr20)) s$kr20 else s$alpha
    rel_lbl <- if((isTRUE(s$is_all_mc) || isTRUE(s$is_all_binary)) && !is.na(s$kr20)) T_("KR20", "KR-20") else T_("CronbachAlpha", "Cronbach Alfa")

    pill <- function(value, class_name = "value-pill-neutral") {
      tags$span(class = paste("value-pill", class_name), ifelse(is.na(value), "NA", round(value, 3)))
    }
    p_class <- if (!is.na(s$avg_p) && s$avg_p >= .4 && s$avg_p <= .8) "value-pill-good" else "value-pill-risk"
    r_class <- if (!is.na(s$avg_r) && s$avg_r >= .3) "value-pill-good" else "value-pill-risk"
    ul27_class <- if (!is.na(s$avg_ul27) && s$avg_ul27 >= .3) "value-pill-good" else "value-pill-risk"

    rows <- list(
      list(T_("StudentsCount", "Ogrenci Sayisi"), s$n_stu),
      list(T_("Median", "Medyan"), round(s$median, 3)),
      list(T_("Mode", "Mod"), round(s$mode, 3)),
      list(T_("Mean", "Aritmetik Ortalama"), round(s$mean, 3)),
      list(T_("SD", "Standart Sapma"), round(s$sd, 3)),
      list(T_("Skewness", "ÃƒÆ’Ã¢â‚¬Â¡arpÃƒâ€Ã‚Â±klÃƒâ€Ã‚Â±k KatsayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"), round(s$skewness, 3)),
      list(T_("Kurtosis", "BasÃƒâ€Ã‚Â±klÃƒâ€Ã‚Â±k KatsayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"), round(s$kurtosis, 3)),
      list(T_("AvgDifficulty", "Ortalama Gucluk (p)"), pill(s$avg_p, p_class)),
      list(T_("AvgDiscrimination", "Ortalama Madde AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i (Corrected Item-Total)"), pill(s$avg_r, r_class)),
      list(T_("AvgDiscriminationUL27", "Ortalama Madde Ayirt Ediciligi (%27 Alt-Ust Gruba Gore)"), pill(s$avg_ul27, ul27_class)),
      list(rel_lbl, round(rel_val, 3)),
      list(T_("MaxWeighted", "Alinabilecek Maksimum Puan"), round(s$max_weighted, 3)),
      list(T_("WeightedExamMean", "Sinav Ortalamasi (katsayili)"), round(mean(s$weighted_total, na.rm = TRUE), 3))
    )

    return(tags$table(
      class = "table table-condensed summary-stat-table",
      tags$thead(
        tags$tr(
          tags$th(T_("Statistic", "Istatistik")),
          tags$th(T_("Value", "Deger"))
        )
      ),
      tags$tbody(
        lapply(rows, function(row) {
          tags$tr(
            tags$td(row[[1]]),
            tags$td(row[[2]])
          )
        })
      )
    ))

    tbl <- tibble::tibble(
      StudentsCount      = s$n_stu,
      Median             = round(s$median, 3),
      Mode               = round(s$mode, 3),
      Mean               = round(s$mean, 3),
      SD                 = round(s$sd, 3),
      AvgDifficulty      = round(s$avg_p, 3),
      AvgDiscrimination  = round(s$avg_r, 3),
      AvgDiscriminationUL27 = round(s$avg_ul27, 3),
      Reliability        = round(rel_val, 3),
      MaxWeighted        = round(s$max_weighted, 3),
      WeightedExamMean   = round(mean(s$weighted_total, na.rm = TRUE), 3)
    )

    # Kept for compatibility with the previous table path.
    labels <- c(
      StudentsCount     = T_("StudentsCount",    "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"),
      Median            = T_("Median",           "Medyan"),
      Mode              = T_("Mode",             "Mod"),
      Mean              = T_("Mean",             "Aritmetik Ortalama"),
      SD                = T_("SD",               "Standart Sapma"),
      AvgDifficulty     = T_("AvgDifficulty",    "Ortalama GÃƒÆ’Ã‚Â¼ÃƒÆ’Ã‚Â§lÃƒÆ’Ã‚Â¼k (p)"),

      AvgDiscrimination = "Ortalama Madde AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i (Corrected Item-Total)",

      AvgDiscriminationUL27 = "Ortalama Madde AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i (%27'lik Alt-ÃƒÆ’Ã…â€œst Gruba GÃƒÆ’Ã‚Â¶re)",

      Reliability       = rel_lbl,
      MaxWeighted       = T_("MaxWeighted",      "AlÃƒâ€Ã‚Â±nabilecek Maksimum Puan"),
      WeightedExamMean  = T_("WeightedExamMean", "SÃƒâ€Ã‚Â±nav OrtalamasÃƒâ€Ã‚Â± (katsayÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â±)")
    )

    # Ãƒâ€Ã‚Â°simleri eÃƒâ€¦Ã…Â¸leÃƒâ€¦Ã…Â¸tir.
    # Ãƒâ€¦Ã‚Âimdilik doÃƒâ€Ã…Â¸rudan atama yapÃƒâ€Ã‚Â±yoruz ki kesin ÃƒÆ’Ã‚Â§alÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸sÃƒâ€Ã‚Â±n.
    names(tbl) <- unname(labels[names(tbl)])
    tbl

  })

  output$avg_p_badge <- renderUI({
    s <- test_summary()
    HTML(paste(T_("AvgP","Ortalama GÃƒÆ’Ã‚Â¼ÃƒÆ’Ã‚Â§lÃƒÆ’Ã‚Â¼k:"), as.character(examly::color_badge(s$avg_p, "p"))))
  })

  output$avg_r_badge <- renderUI({
    s <- test_summary()
    HTML(paste(T_("AvgR","Ortalama AyÃƒâ€Ã‚Â±rt Edicilik:"), as.character(examly::color_badge(s$avg_r, "r"))))
  })

  output$item_stats_table <- renderTable({
    dict()
    s <- test_summary()
    Y <- s$Y; p <- s$p
    item_sd  <- apply(Y, 2, stats::sd, na.rm = TRUE)
    item_var <- item_sd^2

    diff_keys <- vapply(p,     examly::difficulty_label_key,      NA_character_)
    disc_keys <- vapply(s$rjx, examly::discrimination_decision_key, NA_character_)

    diff_labels <- ifelse(is.na(diff_keys), "",
                          vapply(diff_keys, function(k) T_(k, ""), character(1)))
    disc_labels <- ifelse(is.na(disc_keys), "",
                          vapply(disc_keys, function(k) T_(k, ""), character(1)))

    tbl <- tibble::tibble(
      ItemName   = names(Y),
      Difficulty = round(p, 3),
      Discrimination = round(s$rjx, 3),
      UL27Disc   = round(s$ul27, 3),
      SD        = round(item_sd, 3),
      Variance  = round(item_var, 3),
      DifficultyLabel = diff_labels,
      DiscriminationDecision = disc_labels
    )

    colnames(tbl) <- c(
      T_("ItemCol",                 "Soru"),
      T_("DifficultyCol",          "GÃƒÆ’Ã‚Â¼ÃƒÆ’Ã‚Â§lÃƒÆ’Ã‚Â¼k (p)"),
      T_("DiscriminationCol",      "AyÃƒâ€Ã‚Â±rt Edicilik (rjx)"),
      T_("UL27Col",                "AyÃƒâ€Ã‚Â±rt Edicilik (Alt-ÃƒÆ’Ã…â€œst %27)"),
      T_("SD",                     "S.D."),
      T_("VarianceCol",            "Varyans"),
      T_("DifficultyLabelCol",     "Zorluk"),
      T_("DiscDecisionCol",        "Madde AyÃƒâ€Ã‚Â±rt Edicilik Yorumu")
    )
    tbl
  }, align = "c")

  # --- DYNAMIC ITEM COMMENTS UI (FOR I18N) ---
  output$item_comment_ui <- renderUI({
    dict()
    t_disc_title <- T_("Interp.Discrimination", "AyÃƒâ€Ã‚Â±rt Edicilik (Ebel, 1965)")
    t_diff_title <- T_("Interp.Difficulty", "GÃƒÆ’Ã‚Â¼ÃƒÆ’Ã‚Â§lÃƒÆ’Ã‚Â¼k")
    t_val        <- T_("Interp.Value", "DeÃƒâ€Ã…Â¸er")
    t_comment    <- T_("Interp.Comment", "Yorum")

    t_vg  <- T_("Interp.VeryGoodItem", "ÃƒÆ’Ã¢â‚¬Â¡ok iyi madde")
    t_g   <- T_("Interp.GoodItem", "Ãƒâ€Ã‚Â°yi madde (dÃƒÆ’Ã‚Â¼zeltilebilir)")
    t_med <- T_("Interp.MediocreItem", "Orta (dÃƒÆ’Ã‚Â¼zeltilmeli)")
    t_weak<- T_("Interp.WeakItem", "ZayÃƒâ€Ã‚Â±f (ÃƒÆ’Ã‚Â§Ãƒâ€Ã‚Â±karÃƒâ€Ã‚Â±lmalÃƒâ€Ã‚Â±)")

    t_ve  <- T_("Interp.VeryEasy", "ÃƒÆ’Ã¢â‚¬Â¡ok Kolay")
    t_e   <- T_("Interp.Easy", "Kolay")
    t_m   <- T_("Interp.Medium", "Orta")
    t_h   <- T_("Interp.Hard", "Zor")
    t_vh  <- T_("Interp.VeryHard", "ÃƒÆ’Ã¢â‚¬Â¡ok Zor")

    t_note <- T_("Interp.Note", "Tabloda renklendirmeler kolay yorumlanabilirlik aÃƒÆ’Ã‚Â§Ãƒâ€Ã‚Â±sÃƒâ€Ã‚Â±ndan 3 gruba (DÃƒÆ’Ã‚Â¼Ãƒâ€¦Ã…Â¸ÃƒÆ’Ã‚Â¼k/Orta/YÃƒÆ’Ã‚Â¼ksek) indirgenmiÃƒâ€¦Ã…Â¸tir.")

    htmltools::HTML(paste0(
      "<div style='display:flex; gap:20px; flex-wrap:wrap;'>",
      "<div><b>", t_disc_title, ":</b><br>",
      "<table class='table table-condensed table-bordered' style='font-size:12px; width:auto;'>",
      "<tr><th>", t_val, "</th><th>", t_comment, "</th></tr>",
      "<tr><td>> 0.40</td><td>", t_vg, "</td></tr>",
      "<tr><td>0.30 - 0.39</td><td>", t_g, "</td></tr>",
      "<tr><td>0.20 - 0.29</td><td>", t_med, "</td></tr>",
      "<tr><td>< 0.19</td><td>", t_weak, "</td></tr>",
      "</table></div>",

      "<div><b>", t_diff_title, ":</b><br>",
      "<table class='table table-condensed table-bordered' style='font-size:12px; width:auto;'>",
      "<tr><th>", t_val, "</th><th>", t_comment, "</th></tr>",
      "<tr><td>> 0.80</td><td>", t_ve, "</td></tr>",
      "<tr><td>0.60 - 0.80</td><td>", t_e, "</td></tr>",
      "<tr><td>0.40 - 0.60</td><td>", t_m, "</td></tr>",
      "<tr><td>0.20 - 0.40</td><td>", t_h, "</td></tr>",
      "<tr><td>< 0.20</td><td>", t_vh, "</td></tr>",
      "</table></div>",
      "</div>",
      "<p style='margin-top:10px; font-style:italic; color:#666;'>* ", t_note, "</p>"
    ))
  })
  # --------------------------------------------

  mc_raw_norm <- reactive({
    req(mc_items_use()); df <- data_items(); cols <- intersect(names(df), mc_items_use())
    if(length(cols)==0) return(NULL)
    as.data.frame(lapply(cols, function(nm) examly::norm_letter(df[[nm]])), stringsAsFactors=FALSE) %>% setNames(cols)
  })

  compute_distractor <- function(item_name){
    s <- test_summary(); totals <- s$weighted_total
    mlv <- mc_levels(); df_mc <- mc_raw_norm(); keys <- answer_keys()$mc
    if(is.null(df_mc) || !(item_name %in% names(df_mc))) return(NULL)
    x <- df_mc[[item_name]]; key <- toupper(str_trim(keys[[item_name]]))
    n <- length(totals); k <- max(1, floor(n * 0.27)); ord <- order(totals)
    lower_idx <- ord[1:k]; upper_idx <- ord[(n-k+1):n]

    tab_all <- table(factor(x, levels = mlv));     n_all <- as.integer(tab_all)
    pct_all <- if(sum(n_all)==0) rep(NA_real_, length(n_all)) else n_all/sum(n_all)
    tab_low <- table(factor(x[lower_idx], levels = mlv)); n_low <- as.integer(tab_low)
    pct_low <- if(sum(n_low)==0) rep(NA_real_, length(n_low)) else n_low/sum(n_low)
    tab_upp <- table(factor(x[upper_idx], levels = mlv)); n_upp <- as.integer(tab_upp)
    pct_upp <- if(sum(n_upp)==0) rep(NA_real_, length(n_upp)) else n_upp/sum(n_upp)

    discr <- pct_upp - pct_low
    fmt_pct <- function(v){ ifelse(is.na(v), NA_character_, paste0(round(100*v,1),"%")) }

    dict()
    type_correct    <- T_("DistractorType.Correct", "DoÃƒâ€Ã…Â¸ru")
    type_distractor <- T_("DistractorType.Distractor", "ÃƒÆ’Ã¢â‚¬Â¡eldirici")
    col_option   <- T_("DistractorCol.Option",          "SeÃƒÆ’Ã‚Â§enek")
    col_type     <- T_("DistractorCol.Type",            "TÃƒÆ’Ã‚Â¼r")
    col_cnt_all  <- T_("DistractorCol.CountAll",        "Ãƒâ€Ã‚Â°Ãƒâ€¦Ã…Â¸aretleyenlerin SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
    col_pct_all  <- T_("DistractorCol.PctAll",          "Ãƒâ€Ã‚Â°Ãƒâ€¦Ã…Â¸aretleyenlerin YÃƒÆ’Ã‚Â¼zdesi")
    col_cnt_up   <- T_("DistractorCol.CountUpper",      "ÃƒÆ’Ã…â€œst Grupta Ãƒâ€Ã‚Â°Ãƒâ€¦Ã…Â¸aretleyenlerin SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
    col_pct_up   <- T_("DistractorCol.PctUpper",        "ÃƒÆ’Ã…â€œst Grupta Ãƒâ€Ã‚Â°Ãƒâ€¦Ã…Â¸aretleyenler %")
    col_cnt_low  <- T_("DistractorCol.CountLower",      "Alt Grupta Ãƒâ€Ã‚Â°Ãƒâ€¦Ã…Â¸aretleyenlerin SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
    col_pct_low  <- T_("DistractorCol.PctLower",        "Alt Grupta Ãƒâ€Ã‚Â°Ãƒâ€¦Ã…Â¸aretleyenler %")
    col_disc     <- T_("DistractorCol.Discrimination",  "SeÃƒÆ’Ã‚Â§eneÃƒâ€Ã…Â¸in AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i")

    out <- tibble::tibble(
      !!col_option  := mlv,
      !!col_type    := ifelse(mlv == key, type_correct, type_distractor),
      !!col_cnt_all := n_all,
      !!col_pct_all := fmt_pct(pct_all),
      !!col_cnt_up  := n_upp,
      !!col_pct_up  := fmt_pct(pct_upp),
      !!col_cnt_low := n_low,
      !!col_pct_low := fmt_pct(pct_low),
      !!col_disc    := round(discr, 3)
    )
    return(out)
  }
  output$distractor_item_picker <- renderUI({
    dict(); req(mc_items_use())
    selectInput("distr_item", T_("DistractorPickItem","Madde seÃƒÆ’Ã‚Â§in"), choices = mc_items_use())
  })

  output$distractor_highlights_ui <- renderUI({
    dict()
    h <- item_highlights()
    req(h)
    style_box <- "padding: 10px; border: 1px solid #ddd; border-radius: 5px; background: #f9f9f9; margin-bottom: 8px;"
    style_label <- "font-weight: bold; color: #333;"
    style_value_g <- "font-weight: bold; color: #28a745;"
    style_value_r <- "font-weight: bold; color: #dc3545;"
    style_value_b <- "font-weight: bold; color: #6c757d;"
    create_highlight_box <- function(label_key, default_label, value, count, n_total, style_val) {
      if (is.na(value) || is.null(value)) {
        value_str <- T_("NotApplicable", "N/A")
        count_str <- ""
      } else {
        value_str <- value
        pct_str <- if (n_total > 0) sprintf("(%.1f%%)", (count / n_total) * 100) else ""
        count_str <- paste0(" (", count, " ", T_("CountUnit", "kiÃƒâ€¦Ã…Â¸i"), ") ", pct_str)
      }
      tags$div(
        style = style_box,
        tags$span(style = style_label, T_(label_key, default_label), ": "),
        tags$span(style = style_val, value_str),
        tags$span(style="color: #555;", count_str)
      )
    }
    tagList(
      create_highlight_box("MostCorrectItem", "En ÃƒÆ’Ã‚Â§ok doÃƒâ€Ã…Â¸ru yapÃƒâ€Ã‚Â±lan madde", h$most_correct, h$correct_count, h$n_total, style_value_g),
      create_highlight_box("MostWrongItem", "En ÃƒÆ’Ã‚Â§ok yanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸ yapÃƒâ€Ã‚Â±lan madde", h$most_wrong, h$wrong_count, h$n_total, style_value_r),
      create_highlight_box("MostBlankItem", "En ÃƒÆ’Ã‚Â§ok boÃƒâ€¦Ã…Â¸ bÃƒâ€Ã‚Â±rakÃƒâ€Ã‚Â±lan madde", h$most_blank, h$blank_count, h$n_total, style_value_b)
    )
  })

  output$distractor_comment <- renderText({
    dict(); req(input$distr_item)
    tb <- compute_distractor(input$distr_item)
    if (is.null(tb))
      return( T_("DistractorNeedsMCAndKey",
                 "Bu analiz iÃƒÆ’Ã‚Â§in madde ÃƒÆ’Ã‚Â§oktan seÃƒÆ’Ã‚Â§meli olmalÃƒâ€Ã‚Â± ve cevap anahtarÃƒâ€Ã‚Â± girilmiÃƒâ€¦Ã…Â¸ olmalÃƒâ€Ã‚Â±dÃƒâ€Ã‚Â±r.") )
    type_correct    <- T_("DistractorType.Correct", "DoÃƒâ€Ã…Â¸ru")
    type_distractor <- T_("DistractorType.Distractor", "ÃƒÆ’Ã¢â‚¬Â¡eldirici")
    col_type   <- T_("DistractorCol.Type", "TÃƒÆ’Ã‚Â¼r")
    col_disc   <- T_("DistractorCol.Discrimination", "SeÃƒÆ’Ã‚Â§eneÃƒâ€Ã…Â¸in AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i")
    col_option <- T_("DistractorCol.Option", "SeÃƒÆ’Ã‚Â§enek")
    key_row <- tb %>% dplyr::filter(.data[[col_type]] == type_correct)
    ok <- nrow(key_row) > 0 && !is.na(key_row[[col_disc]][1]) && key_row[[col_disc]][1] > 0
    bads <- tb %>% dplyr::filter(.data[[col_type]] == type_distractor, .data[[col_disc]] > 0)
    msg1 <- if (ok) T_("DistractorMsg_OK",
                       "DoÃƒâ€Ã…Â¸ru seÃƒÆ’Ã‚Â§enek baÃƒâ€¦Ã…Â¸arÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â± ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸renciler tarafÃƒâ€Ã‚Â±ndan daha ÃƒÆ’Ã‚Â§ok iÃƒâ€¦Ã…Â¸aretlenmiÃƒâ€¦Ã…Â¸.")
    else     T_("DistractorMsg_NotOK",
                "DoÃƒâ€Ã…Â¸ru seÃƒÆ’Ã‚Â§enek baÃƒâ€¦Ã…Â¸arÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â± ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸rencilerde daha ÃƒÆ’Ã‚Â§ok gÃƒÆ’Ã‚Â¶rÃƒÆ’Ã‚Â¼nmÃƒÆ’Ã‚Â¼yor.")
    msg2 <- if (nrow(bads) == 0)
      paste0("\n", T_("DistractorMsg_OkDistractors",
                      "ÃƒÆ’Ã¢â‚¬Â¡eldiriciler baÃƒâ€¦Ã…Â¸arÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â± dÃƒÆ’Ã‚Â¼Ãƒâ€¦Ã…Â¸ÃƒÆ’Ã‚Â¼k ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸renciler tarafÃƒâ€Ã‚Â±ndan daha ÃƒÆ’Ã‚Â§ok seÃƒÆ’Ã‚Â§iliyor."))
    else
      paste0("\n",
             T_("DistractorMsg_BadDistractorsPrefix",
                "BaÃƒâ€¦Ã…Â¸arÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â± ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸rencilerde gÃƒÆ’Ã‚Â¶rece daha ÃƒÆ’Ã‚Â§ok seÃƒÆ’Ã‚Â§ilen ÃƒÆ’Ã‚Â§eldiriciler: "),
             paste0(bads[[col_option]], collapse = ", "), ".")
    paste(msg1, msg2)
  })

  output$distractor_table <- renderTable({ req(input$distr_item); compute_distractor(input$distr_item) }, align="c")

  student_item_comment <- function(sc_row, lc_row, w, mc_names, tf_names, bin_names, lc_names, outcomes = NULL){
    scn <- suppressWarnings(as.numeric(sc_row)); names(scn) <- names(sc_row)
    scn[is.na(scn)] <- NA_real_
    zero_mctf <- names(scn)[names(scn) %in% union(mc_names, tf_names) & !is.na(scn) & scn == 0]
    zero_bin  <- names(scn)[names(scn) %in% bin_names                 & !is.na(scn) & scn == 0]
    lcn <- suppressWarnings(as.numeric(lc_row)); names(lcn) <- names(lc_row)
    w0  <- suppressWarnings(as.numeric(w[names(lc_row)]));   names(w0)  <- names(lc_row)
    low_lc <- names(lcn)[names(lcn) %in% lc_names & !is.na(lcn) & !is.na(w0) & (lcn < 0.5 * w0)]
    if(length(zero_mctf) == 0 && length(zero_bin) == 0 && length(low_lc) == 0) return("")
    weak_items <- sort(unique(c(zero_mctf, zero_bin, low_lc)))
    if (!is.null(outcomes)) {
      mapped <- outcomes[weak_items]
      mapped <- mapped[!is.na(mapped) & mapped != ""]
      if (length(mapped) > 0) weak_items <- sort(unique(unname(mapped)))
    }
    paste0(" ", paste0(weak_items, collapse = ", "), ".")
  }

  output$student_table_html <- renderUI({
    s <- test_summary()
    dict()
    if (is.null(s)) return(tags$p( T_("SummaryNotReady",
                                      "ÃƒÆ’Ã¢â‚¬â€œzet hesaplanamadÃƒâ€Ã‚Â±. LÃƒÆ’Ã‚Â¼tfen veri ve anahtarlarÃƒâ€Ã‚Â± kontrol edin.") ))

    sc_all <- s$sc_bin; lc_all <- s$lc_raw
    # Hata buradaydÃƒâ€Ã‚Â±: bnN tanÃƒâ€Ã‚Â±mlÃƒâ€Ã‚Â±ydÃƒâ€Ã‚Â± ama aÃƒâ€¦Ã…Â¸aÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±da binN ÃƒÆ’Ã‚Â§aÃƒâ€Ã…Â¸rÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â±yordu.
    mcN <- mc_items_use(); tfN <- tf_items_use(); lcN <- lc_items_use(); bnN <- bin_items_use(); w <- s$weights
    outcomes <- item_outcomes()
    has_mc <- length(mcN) > 0; has_tf <- length(tfN) > 0; has_bn <- length(bnN) > 0; has_lc <- length(lcN) > 0

    cnt_mc <- if(has_mc) examly::student_counts(sc_all[, intersect(colnames(sc_all), mcN), drop=FALSE]) else NULL
    cnt_tf <- if(has_tf) examly::student_counts(sc_all[, intersect(colnames(sc_all), tfN), drop=FALSE]) else NULL
    cnt_bn <- if(has_bn) examly::student_counts(sc_all[, intersect(colnames(sc_all), bnN), drop=FALSE]) else NULL
    lc_sum <- if(has_lc) rowSums(lc_all[, intersect(colnames(lc_all), lcN), drop=FALSE], na.rm=TRUE) else rep(0, nrow(sc_all))

    yorumlar <- purrr::map_chr(seq_len(nrow(s$sc_bin)), function(i){
      row_mctf <- s$sc_bin[i,,drop=TRUE]; row_lc <- s$lc_raw[i,,drop=TRUE]
      # DÃƒÆ’Ã…â€œZELTME: binN -> bnN olarak deÃƒâ€Ã…Â¸iÃƒâ€¦Ã…Â¸tirildi
      student_item_comment(row_mctf, row_lc, w, mcN, tfN, bnN, lcN)
    })

    tbl_style <- "width:100%; border-collapse:collapse;"
    th_style <- "border:1px solid #e5e7eb; padding:8px; background:#f9fafb; text-align:center; font-weight:600;"
    td_style <- "border:1px solid #e5e7eb; padding:6px; text-align:center;"
    name_style <- "border:1px solid #e5e7eb; padding:6px; text-align:left;"

    dict()
    lbl_student   <- T_("StudentsColHeader", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci")
    lbl_mc_group  <- T_("MCGroupHeader", "ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli Maddeler")
    lbl_tf_group  <- T_("TFGroupHeader", "D/Y Maddeler")
    lbl_bin_group <- T_("BINGroupHeader", "1-0 Maddeler")
    lbl_lc_group  <- T_("LCGroupHeader", "Uzun CevaplÃƒâ€Ã‚Â± (AlÃƒâ€Ã‚Â±nan Puan)")
    lbl_correct   <- T_("Correct", "DoÃƒâ€Ã…Â¸ru"); lbl_wrong     <- T_("Wrong",   "YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸"); lbl_blank     <- T_("Blank",   "BoÃƒâ€¦Ã…Â¸")
    lbl_wscore    <- T_("WeightedScore", "KatsayÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â± Puan"); lbl_weak      <- T_("StudentWeakItems", "KazanÃƒâ€Ã‚Â±mlara Ait Eksiklik Olabilecek Maddeler")

    header1_elements <- list(tags$th(style=th_style, rowspan=2, lbl_student))
    if (has_mc) header1_elements <- c(header1_elements, list(tags$th(style=th_style, colspan=3, lbl_mc_group)))
    if (has_tf) header1_elements <- c(header1_elements, list(tags$th(style=th_style, colspan=3, lbl_tf_group)))
    if (has_bn) header1_elements <- c(header1_elements, list(tags$th(style=th_style, colspan=3, lbl_bin_group)))
    if (has_lc) header1_elements <- c(header1_elements, list(tags$th(style=th_style, rowspan=2, lbl_lc_group)))
    header1_elements <- c(header1_elements, list(tags$th(style=th_style, rowspan=2, lbl_wscore), tags$th(style=th_style, rowspan=2, lbl_weak)))
    header1 <- do.call(tags$tr, header1_elements)

    header2_elements <- list()
    if (has_mc) { header2_elements <- c(header2_elements, list(tags$th(style=th_style, lbl_correct), tags$th(style=th_style, lbl_wrong), tags$th(style=th_style, lbl_blank))) }
    if (has_tf) { header2_elements <- c(header2_elements, list(tags$th(style=th_style, lbl_correct), tags$th(style=th_style, lbl_wrong), tags$th(style=th_style, lbl_blank))) }
    if (has_bn){ header2_elements <- c(header2_elements, list(tags$th(style=th_style, lbl_correct), tags$th(style=th_style, lbl_wrong), tags$th(style=th_style, lbl_blank))) }
    header2 <- if (length(header2_elements) > 0) do.call(tags$tr, header2_elements) else NULL

    rows <- lapply(seq_len(nrow(sc_all)), function(i){
      row_elements <- list(tags$td(style=name_style, student_names()[i]))
      if (has_mc) { row_elements <- c(row_elements, list(tags$td(style=td_style, cnt_mc$Dogru[i]), tags$td(style=td_style, cnt_mc$Yanlis[i]), tags$td(style=td_style, cnt_mc$Bos[i]))) }
      if (has_tf) { row_elements <- c(row_elements, list(tags$td(style=td_style, cnt_tf$Dogru[i]), tags$td(style=td_style, cnt_tf$Yanlis[i]), tags$td(style=td_style, cnt_tf$Bos[i]))) }
      if (has_bn) { row_elements <- c(row_elements, list(tags$td(style=td_style, cnt_bn$Dogru[i]), tags$td(style=td_style, cnt_bn$Yanlis[i]), tags$td(style=td_style, cnt_bn$Bos[i]))) }
      if (has_lc) { row_elements <- c(row_elements, list(tags$td(style=td_style, round(lc_sum[i],3)))) }
      row_elements <- c(row_elements, list(tags$td(style=td_style, round(s$weighted_total[i],3)), tags$td(style=paste0(td_style, " text-align:left;"), yorumlar[i])))
      do.call(tags$tr, row_elements)
    })
    tags$table(style=tbl_style, tags$thead(header1, header2), tags$tbody(rows))
  })

  output$student_summary_text_ui <- renderUI({
    input$lang; s_proc <- student_scores_processed(); req(s_proc); dict()
    pass_text <- T_("StudentsPassed50", "AÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±rlÃƒâ€Ã‚Â±klÃƒâ€Ã‚Â± puana gÃƒÆ’Ã‚Â¶re 50 ve ÃƒÆ’Ã‚Â¼zeri alan ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸renci sayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
    pct_str <- ""
    if (s_proc$n_total > 0) {
      pct_pass <- (s_proc$count_pass_50 / s_proc$n_total) * 100
      pct_str <- sprintf(" (%%%.1f)", pct_pass)
    }
    tags$p(style = "font-size: 16px;", strong(paste0(pass_text, ": ", s_proc$count_pass_50, pct_str)), " (", T_("TotalStudents", "Toplam"), ": ", s_proc$n_total, ")")
  })

  item_failure_rank <- reactive({
    s <- test_summary()
    req(s)
    outcomes <- item_outcomes()
    all_items <- colnames(s$sc_bin)
    counts <- vapply(all_items, function(nm) {
      if (nm %in% lc_items_use()) {
        v <- suppressWarnings(as.numeric(s$lc_raw[[nm]]))
        mx <- suppressWarnings(as.numeric(s$weights[[nm]]))
        if (is.na(mx)) return(0L)
        sum(!is.na(v) & v < 0.5 * mx)
      } else {
        v <- suppressWarnings(as.numeric(s$sc_bin[[nm]]))
        sum(!is.na(v) & v == 0)
      }
    }, integer(1))
    n_total <- nrow(s$sc_bin)
    tibble::tibble(
      Item = all_items,
      Outcome = unname(ifelse(!is.na(outcomes[all_items]) & outcomes[all_items] != "", outcomes[all_items], "")),
      FailureCount = as.integer(counts),
      FailurePct = if (n_total > 0) round((counts / n_total) * 100, 1) else NA_real_
    ) %>%
      dplyr::arrange(dplyr::desc(.data$FailureCount), .data$Item)
  })

  output$item_failure_rank_ui <- renderUI({
    dict()
    rank_tbl <- item_failure_rank()
    req(nrow(rank_tbl) > 0)
    tagList(
      h4(class = "failure-rank-title", T_("MostFailedItemsRank", "En ÃƒÆ’Ã¢â‚¬Â¡ok YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸ YapÃƒâ€Ã‚Â±lan Madde SÃƒâ€Ã‚Â±ralamasÃƒâ€Ã‚Â±")),
      tags$div(class = "preview-scroll", tableOutput("item_failure_rank_table"))
    )
  })

  output$item_failure_rank_table <- renderTable({
    dict()
    rank_tbl <- item_failure_rank()
    colnames(rank_tbl) <- c(
      T_("ItemCol", "Madde"),
      T_("LearningOutcome", "KazanÃƒâ€Ã‚Â±m"),
      T_("FailureCount", "YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸/DÃƒÆ’Ã‚Â¼Ãƒâ€¦Ã…Â¸ÃƒÆ’Ã‚Â¼k Puan SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"),
      T_("FailurePct", "Oran (%)")
    )
    rank_tbl
  }, align = "c")

  output$student_grade_plot <- renderPlot({
    s_proc <- student_scores_processed(); req(s_proc, nrow(s_proc$df_grades) > 0); dict()
    df_plot <- s_proc$df_grades %>%
      dplyr::group_by(Grades) %>%
      dplyr::summarise(Count = n(), .groups = 'drop') %>%
      tidyr::complete(Grades = factor(s_proc$grades_i18n, levels = s_proc$grades_i18n), fill = list(Count = 0)) %>%
      dplyr::mutate(
        Total = sum(Count),
        Pct = ifelse(Total == 0, 0, Count / Total),
        Label = ifelse(Count == 0, "", sprintf("%d\n(%.1f%%)", Count, Pct * 100))
      )
    ggplot(df_plot, aes(x = Grades, y = Count, fill = Grades)) +
      geom_col(width = 0.64, show.legend = FALSE, color = "white", linewidth = 0.8) +
      geom_text(aes(label = Label), vjust = -0.45, size = 4.2, lineheight = 0.9, fontface = "bold", color = "#162033") +
      scale_fill_manual(values = c("#dc2626", "#f97316", "#f59e0b", "#22c55e", "#16a34a")[seq_along(levels(df_plot$Grades))]) +
      scale_y_continuous(expand = expansion(mult = c(0, 0.15))) +
      labs(title = T_("GradeDistributionTitle", "Puan Kategorilerine GÃƒÆ’Ã‚Â¶re ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci DaÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±lÃƒâ€Ã‚Â±mÃƒâ€Ã‚Â±"), x = T_("GradeCategory", "Puan Kategorisi"), y = T_("StudentCount", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")) +
      theme_minimal(base_size = 14) +
      theme(
        plot.background = element_rect(fill = "#f8fafc", color = NA),
        panel.background = element_rect(fill = "#f8fafc", color = NA),
        panel.grid.major.x = element_blank(),
        panel.grid.minor = element_blank(),
        panel.grid.major.y = element_line(color = "#dbe4ef", linewidth = 0.35),
        axis.text.x = element_text(angle = 0, hjust = 0.5, size = 12, face = "bold", color = "#334155"),
        axis.text.y = element_text(color = "#64748b"),
        plot.title = element_text(size = 18, face = "bold", color = "#162033", margin = margin(b = 14)),
        axis.title = element_text(size = 13, face = "bold", color = "#475569")
      )
  })

  make_student_df <- function(s, mcN, tfN, binN, lcN, names_vec){
    dict()
    sc_all <- s$sc_bin; lc_all <- s$lc_raw
    has_mc <- length(mcN)  > 0; has_tf <- length(tfN)  > 0; has_bn <- length(binN) > 0; has_lc <- length(lcN)  > 0
    col_student <- T_("StudentsColHeader", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci")
    short_mc    <- T_("MCShort",  "ÃƒÆ’Ã¢â‚¬Â¡S"); short_tf    <- T_("TFShort",  "D/Y"); short_bin   <- T_("BINShort", "1-0"); lc_sum_lbl  <- T_("LCSumShort","Uzun Cev. Toplam")
    lbl_correct <- T_("Correct", "DoÃƒâ€Ã…Â¸ru"); lbl_wrong   <- T_("Wrong",   "YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸"); lbl_blank   <- T_("Blank",   "BoÃƒâ€¦Ã…Â¸")
    col_wscore  <- T_("WeightedScore",   "KatsayÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â± Puan"); col_weak    <- T_("StudentWeakItems","KazanÃƒâ€Ã‚Â±mlara Ait Eksiklik Olabilecek Maddeler")

    outcomes <- item_outcomes()
    result_list <- list(); result_list[[col_student]] <- unname(names_vec)
    if (has_mc) {
      cnt_mc <- examly::student_counts(sc_all[, intersect(colnames(sc_all), mcN), drop=FALSE])
      result_list[[paste(short_mc, lbl_correct)]] <- unname(cnt_mc$Dogru)
      result_list[[paste(short_mc, lbl_wrong)  ]] <- unname(cnt_mc$Yanlis)
      result_list[[paste(short_mc, lbl_blank)  ]] <- unname(cnt_mc$Bos)
    }
    if (has_tf) {
      cnt_tf <- examly::student_counts(sc_all[, intersect(colnames(sc_all), tfN), drop=FALSE])
      result_list[[paste(short_tf, lbl_correct)]] <- unname(cnt_tf$Dogru)
      result_list[[paste(short_tf, lbl_wrong)  ]] <- unname(cnt_tf$Yanlis)
      result_list[[paste(short_tf, lbl_blank)  ]] <- unname(cnt_tf$Bos)
    }
    if (has_bn) {
      cnt_bn <- examly::student_counts(sc_all[, intersect(colnames(sc_all), binN), drop=FALSE])
      result_list[[paste(short_bin, lbl_correct)]] <- unname(cnt_bn$Dogru)
      result_list[[paste(short_bin, lbl_wrong)  ]] <- unname(cnt_bn$Yanlis)
      result_list[[paste(short_bin, lbl_blank)  ]] <- unname(cnt_bn$Bos)
    }
    if (has_lc) {
      lc_sum <- rowSums(lc_all[, intersect(colnames(lc_all), lcN), drop=FALSE], na.rm=TRUE)
      result_list[[lc_sum_lbl]] <- unname(round(lc_sum, 3))
    }
    yorumlar <- purrr::map_chr(seq_len(nrow(s$sc_bin)), function(i){
      row_mctf <- s$sc_bin[i,,drop=TRUE]; row_lc <- s$lc_raw[i,,drop=TRUE]
      student_item_comment(row_mctf, row_lc, s$weights, mcN, tfN, binN, lcN)
    })
    result_list[[col_wscore]] <- unname(round(s$weighted_total, 3))
    result_list[[col_weak]]   <- unname(yorumlar)
    as_tibble(result_list)
  }

  add_multiheader_ft <- function(ft){
    dict()
    df_cols <- names(ft$body$dataset)
    short_mc   <- T_("MCShort",  "ÃƒÆ’Ã¢â‚¬Â¡S"); short_tf   <- T_("TFShort",  "D/Y"); short_bin  <- T_("BINShort", "1-0"); lc_sum_lbl <- T_("LCSumShort","Uzun Cev. Toplam")
    g_mc <- T_("MCGroupHeader", "ÃƒÆ’Ã¢â‚¬Â¡oktan SeÃƒÆ’Ã‚Â§meli Maddeler"); g_tf <- T_("TFGroupHeader", "D/Y Maddeler"); g_bn <- T_("BINGroupHeader","1-0 Maddeler"); g_lc <- T_("LCGroupHeader", "Uzun CevaplÃƒâ€Ã‚Â± (AlÃƒâ€Ã‚Â±nan Puan)")
    lbl_correct <- T_("Correct","DoÃƒâ€Ã…Â¸ru"); lbl_wrong   <- T_("Wrong","YanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸"); lbl_blank   <- T_("Blank","BoÃƒâ€¦Ã…Â¸")

    has_mc <- any(df_cols %in% paste(short_mc, c(lbl_correct, lbl_wrong, lbl_blank)))
    has_tf <- any(df_cols %in% paste(short_tf, c(lbl_correct, lbl_wrong, lbl_blank)))
    has_bn <- any(df_cols %in% paste(short_bin, c(lbl_correct, lbl_wrong, lbl_blank)))
    has_lc <- lc_sum_lbl %in% df_cols

    header_values   <- list(""); header_colwidth <- list(1)
    if (has_mc){ header_values <- c(header_values, g_mc); header_colwidth <- c(header_colwidth, 3) }
    if (has_tf){ header_values <- c(header_values, g_tf); header_colwidth <- c(header_colwidth, 3) }
    if (has_bn){ header_values <- c(header_values, g_bn); header_colwidth <- c(header_colwidth, 3) }
    if (has_lc){ header_values <- c(header_values, "") ; header_colwidth <- c(header_colwidth, 1) }
    header_values   <- c(header_values, "", ""); header_colwidth <- c(header_colwidth, 1, 1)

    ft %>%
      add_header_row(values = unlist(header_values), colwidths = unlist(header_colwidth)) %>%
      merge_h(part = "header") %>% align(align = "center", part = "all") %>% theme_vanilla() %>% autofit()
  }

  make_student_docx <- function(file, title, s, mcN, tfN, binN, lcN, names_vec){
    df <- make_student_df(s, mcN, tfN, binN, lcN, names_vec)
    doc <- read_docx()
    doc <- body_add_par(doc, title, style = "heading 1")
    doc <- add_report_metadata(doc)
    ft <- flextable(df); ft <- add_multiheader_ft(ft)
    doc <- body_add_flextable(doc, ft)
    print(doc, target = file)
  }
  make_summary_docx <- function(file,title,s, mcN, tfN, binN, lcN, names_vec, test_tbl,item_tbl, dlist){
    doc<-read_docx()
    doc<-body_add_par(doc,title,style="heading 1")
    doc<-add_report_metadata(doc)
    doc<-body_add_flextable(doc,flextable(test_tbl) %>% autofit() %>% fit_to_width(6.5))
    body_add_par(doc, T_("ReportHead.ItemStats", "Madde Ãƒâ€Ã‚Â°statistikleri"), style="heading 2")
    body_add_par(doc, T_("ReportHead.StudentResults", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SonuÃƒÆ’Ã‚Â§larÃƒâ€Ã‚Â±"), style="heading 2")
    ft_students <- flextable(make_student_df(s, mcN, tfN, binN, lcN, names_vec)) %>% add_multiheader_ft()
    doc <- body_add_flextable(doc, ft_students)

    if(length(dlist)>0){
      body_add_par(doc, T_("ReportHead.Distractor", "ÃƒÆ’Ã¢â‚¬Â¡eldirici Analizi"), style = "heading 2")
      for(nm in names(dlist)){
        body_add_par(doc, nm, style = "heading 3")
        doc <- body_add_flextable(doc, flextable(dlist[[nm]]) %>% autofit() %>% fit_to_width(6.5))
      }
    }
    print(doc,target=file)
  }

  output$dl_rapor_ogrenci <- downloadHandler(
    filename = function() { paste0("ogrenci_sonuclari_", Sys.Date(), ".docx") },
    content = function(file) {
      dict()
      tryCatch({
        s <- test_summary()
        if (is.null(s)) {
          stop(T_("ReportErrorNullSummary", "Rapor oluÃƒâ€¦Ã…Â¸turulamadÃƒâ€Ã‚Â±: Test ÃƒÆ’Ã‚Â¶zeti boÃƒâ€¦Ã…Â¸."))
        }
        mcN <- mc_items_use(); tfN <- tf_items_use(); binN <- bin_items_use(); lcN <- lc_items_use()
        names_vec <- student_names()
        if (is.null(names_vec) || length(names_vec) == 0) {
          stop(T_("ReportErrorNullNames", "Rapor oluÃƒâ€¦Ã…Â¸turulamadÃƒâ€Ã‚Â±: ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci isimleri boÃƒâ€¦Ã…Â¸."))
        }

        df_students <- tryCatch({
          make_student_df(s, mcN, tfN, binN, lcN, names_vec)
        }, error = function(e_df) {
          stop(paste(T_("ReportError", "Rapor oluÃƒâ€¦Ã…Â¸turulamadÃƒâ€Ã‚Â±. Hata MesajÃƒâ€Ã‚Â±: "), e_df$message))
        })

        if (is.null(df_students) || !is.data.frame(df_students) || nrow(df_students) == 0) {
          stop(T_("ReportError", "Rapor oluÃƒâ€¦Ã…Â¸turulamadÃƒâ€Ã‚Â±. Hata MesajÃƒâ€Ã‚Â±: "), "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci veri ÃƒÆ’Ã‚Â§erÃƒÆ’Ã‚Â§evesi boÃƒâ€¦Ã…Â¸/geÃƒÆ’Ã‚Â§ersiz.")
        }

        doc <- read_docx()
        doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
        doc <- body_add_par(
          doc,
          paste0(exam_title(), " - ", T_("ReportHead.StudentResults","ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SonuÃƒÆ’Ã‚Â§larÃƒâ€Ã‚Â±")),
          style = "heading 1"
        )
        doc <- add_report_metadata(doc)

        ft_students <- flextable(df_students)
        ft_students <- add_multiheader_ft(ft_students)
        ft_students <- fit_to_width(ft_students, max_width = 9.5)

        doc <- body_add_flextable(doc, ft_students)

        s_proc <- student_scores_processed()
        if (!is.null(s_proc)) {
          pass_text <- T_("StudentsPassed50", "AÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±rlÃƒâ€Ã‚Â±klÃƒâ€Ã‚Â± puana gÃƒÆ’Ã‚Â¶re 50 ve ÃƒÆ’Ã‚Â¼zeri alan ÃƒÆ’Ã‚Â¶Ãƒâ€Ã…Â¸renci sayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
          pct_str <- ""
          if (s_proc$n_total > 0) {
            pct_pass <- (s_proc$count_pass_50 / s_proc$n_total) * 100
            pct_str <- sprintf(" (%%%.1f)", pct_pass)
          }
          summary_p_text <- paste0(
            pass_text, ": ", s_proc$count_pass_50, pct_str,
            " (", T_("TotalStudents", "Toplam"), ": ", s_proc$n_total, ")"
          )

          doc <- body_add_par(doc, "")
          doc <- body_add_par(doc, summary_p_text, style = "Normal")
        }

        if (!is.null(s_proc) && nrow(s_proc$df_grades) > 0) {

          df_plot <- s_proc$df_grades %>%
            dplyr::group_by(Grades) %>%
            dplyr::summarise(Count = n(), .groups = 'drop') %>%
            tidyr::complete(Grades = factor(s_proc$grades_i18n, levels = s_proc$grades_i18n),
                            fill = list(Count = 0)) %>%
            dplyr::mutate(
              Total = sum(Count),
              Pct = ifelse(Total == 0, 0, Count / Total),
              Label = ifelse(Count == 0, "", sprintf("%d\n(%.1f%%)", Count, Pct * 100))
            )


          p <- ggplot(df_plot, aes(x = Grades, y = Count, fill = Grades)) +
            geom_bar(stat = "identity", show.legend = FALSE) +
            geom_text(aes(label = Label), vjust = -0.5, size = 3.5, lineheight = 0.9) +
            scale_y_continuous(expand = expansion(mult = c(0, 0.15))) +
            labs(
              title = T_("GradeDistributionTitle", "Puan Kategorilerine GÃƒÆ’Ã‚Â¶re ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci DaÃƒâ€Ã…Â¸Ãƒâ€Ã‚Â±lÃƒâ€Ã‚Â±mÃƒâ€Ã‚Â±"),
              x = T_("GradeCategory", "Puan Kategorisi"),
              y = T_("StudentCount", "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±")
            ) +
            theme_minimal(base_size = 11) +
            theme(axis.text.x = element_text(angle = 0, hjust = 0.5))

          temp_plot_file <- tempfile(fileext = ".png")
          ggsave(temp_plot_file, plot = p, width = 8, height = 5, dpi = 150)

          doc <- body_add_par(doc, "")
          doc <- body_add_img(doc, src = temp_plot_file, width = 8, height = 5)

          unlink(temp_plot_file)
        }

        print(doc, target = file)

      }, error = function(e) {
        msg <- paste0(T_("ReportError", "Rapor oluÃƒâ€¦Ã…Â¸turulamadÃƒâ€Ã‚Â±. Hata MesajÃƒâ€Ã‚Â±: "), conditionMessage(e))
        con <- file(file, open = "w", encoding = "UTF-8")
        on.exit(close(con), add = TRUE)
        writeLines(msg, con = con)
      })
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )

  output$dl_rapor_madde <- downloadHandler(
    filename = function() { paste0("madde_istatistikleri_", Sys.Date(), ".docx") },
    content = function(file) {
      dict()
      s <- test_summary(); req(s)
      Y <- s$Y; p <- s$p

      item_sd  <- apply(Y, 2, stats::sd, na.rm = TRUE)
      item_var <- item_sd^2
      reliab <- sapply(seq_len(ncol(Y)), function(j){
        tryCatch(stats::cor(Y[[j]], s$weighted_total, use = "pairwise.complete.obs"),
                 error = function(e) NA_real_)
      })

      diff_keys <- vapply(p,     examly::difficulty_label_key,        NA_character_)
      disc_keys <- vapply(s$rjx, examly::discrimination_decision_key, NA_character_)

      diff_labels <- ifelse(is.na(diff_keys), "",
                            vapply(diff_keys, function(k) T_(k, ""), character(1)))
      disc_labels <- ifelse(is.na(disc_keys), "",
                            vapply(disc_keys, function(k) T_(k, ""), character(1)))

      item_tbl <- tibble::tibble(
        ItemName               = names(Y),
        Difficulty             = round(p, 3),
        Discrimination         = round(s$rjx, 3),
        UL27Disc               = round(s$ul27, 3), # Rapor ciktiya da eklendi
        SD                     = round(item_sd, 3),
        Variance               = round(item_var, 3),
        ItemReliabilityIndex   = round(reliab, 3),
        DifficultyLabel        = diff_labels,
        DiscriminationDecision = disc_labels
      )

      colnames(item_tbl) <- c(
        T_("ItemCol",            "Soru"),
        T_("DifficultyCol",      "GÃƒÆ’Ã‚Â¼ÃƒÆ’Ã‚Â§lÃƒÆ’Ã‚Â¼k"),
        T_("DiscriminationCol",  "AyÃƒâ€Ã‚Â±rt Edicilik (rjx)"),
        T_("UL27Col",            "AyÃƒâ€Ã‚Â±rt Edicilik (Alt-ÃƒÆ’Ã…â€œst %27)"),
        T_("SD",                 "S.D."),
        T_("VarianceCol",        "Varyans"),
        T_("ItemRelIndexCol",    "Madde GÃƒÆ’Ã‚Â¼venirlik Ãƒâ€Ã‚Â°ndeksi"),
        T_("DifficultyLabelCol", "Zorluk"),
        T_("DiscDecisionCol",    "Madde AyÃƒâ€Ã‚Â±rt Edicilik Yorumu")
      )

      doc <- read_docx()
      doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
      doc <- body_add_par(
        doc,
        paste0(exam_title(), " - ", T_("ReportHead.ItemStats","Madde Ãƒâ€Ã‚Â°statistikleri")),
        style = "heading 1"
      )
      doc <- add_report_metadata(doc)
      ft_items <- flextable::flextable(item_tbl) %>%
        flextable::autofit() %>%
        flextable::fit_to_width(max_width = 9.5)
      doc <- body_add_flextable(doc, ft_items)
      h <- item_highlights()
      if (!is.null(h) && h$n_total > 0) {

        doc <- body_add_par(doc, "")
        doc <- body_add_par(doc, T_("DistractorHighlightsHeader", "Madde ÃƒÆ’Ã¢â‚¬â€œzeti"), style = "heading 2")

        create_report_line <- function(label_key, default_label, value, count, n_total) {
          label <- T_(label_key, default_label)
          if (is.na(value) || is.null(value)) {
            value_str <- T_("NotApplicable", "N/A")
            count_str <- ""
          } else {
            value_str <- value
            pct_str <- sprintf("(%.1f%%)", (count / n_total) * 100)
            count_str <- paste0(" (", count, " ", T_("CountUnit", "kiÃƒâ€¦Ã…Â¸i"), ") ", pct_str)
          }
          paste0(label, ": ", value_str, count_str)
        }

        line1 <- create_report_line("MostCorrectItem", "En ÃƒÆ’Ã‚Â§ok doÃƒâ€Ã…Â¸ru yapÃƒâ€Ã‚Â±lan madde", h$most_correct, h$correct_count, h$n_total)
        line2 <- create_report_line("MostWrongItem", "En ÃƒÆ’Ã‚Â§ok yanlÃƒâ€Ã‚Â±Ãƒâ€¦Ã…Â¸ yapÃƒâ€Ã‚Â±lan madde", h$most_wrong, h$wrong_count, h$n_total)
        line3 <- create_report_line("MostBlankItem", "En ÃƒÆ’Ã‚Â§ok boÃƒâ€¦Ã…Â¸ bÃƒâ€Ã‚Â±rakÃƒâ€Ã‚Â±lan madde", h$most_blank, h$blank_count, h$n_total)

        doc <- body_add_par(doc, line1, style = "Normal")
        doc <- body_add_par(doc, line2, style = "Normal")
        doc <- body_add_par(doc, line3, style = "Normal")
      }
      print(doc, target = file)
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )

  output$dl_rapor_celdirici <- downloadHandler(
    filename = function() { paste0("celdirici_analizi_", Sys.Date(), ".docx") },
    content = function(file) {
      s <- test_summary(); req(s)
      mcN <- mc_items_use()
      doc <- read_docx()
      doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
      dict()
      doc <- body_add_par(
        doc,
        paste0(exam_title(), " - ", T_("ReportHead.Distractor","ÃƒÆ’Ã¢â‚¬Â¡eldirici Analizi")),
        style = "heading 1"
      )
      doc <- add_report_metadata(doc)

      if (length(mcN) == 0) {
        doc <- body_add_par(doc, T_("NoMCForDistractor","Analiz edilecek ÃƒÆ’Ã‚Â§oktan seÃƒÆ’Ã‚Â§meli madde yok."))
      } else {
        for(nm in mcN){
          tb <- compute_distractor(nm)
          if(!is.null(tb)){
            doc <- body_add_par(
              doc,
              paste0(T_("ItemCol", "Madde"), ": ", nm),
              style = "heading 2"
            )
            ft_dist <- flextable(tb) %>% autofit() %>% fit_to_width(max_width = 9.5)
            doc <- body_add_flextable(doc, ft_dist)
            doc <- body_add_par(doc, "")
          }
        }
      }
      print(doc, target = file)
    }
  )

  output$dl_rapor_ozet <- downloadHandler(
    filename = function() { paste0("test_ozeti_", Sys.Date(), ".docx") },
    content = function(file) {
      dict()
      s <- test_summary(); req(s)

      # Binary veri kontrolune gore KR-20 veya Alpha secimi
      rel_val <- if((isTRUE(s$is_all_mc) || isTRUE(s$is_all_binary)) && !is.na(s$kr20)) s$kr20 else s$alpha
      rel_lbl <- if((isTRUE(s$is_all_mc) || isTRUE(s$is_all_binary)) && !is.na(s$kr20)) T_("KR20", "KR-20") else T_("CronbachAlpha", "Cronbach Alfa")

      row_labels <- c(
        StudentsCount     = T_("StudentsCount",     "ÃƒÆ’Ã¢â‚¬â€œÃƒâ€Ã…Â¸renci SayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"),
        Median            = T_("Median",            "Medyan"),
        Mode              = T_("Mode",              "Mod"),
        Mean              = T_("Mean",              "Aritmetik Ortalama"),
        SD                = T_("SD",                "Standart Sapma"),
        Skewness          = T_("Skewness",          "ÃƒÆ’Ã¢â‚¬Â¡arpÃƒâ€Ã‚Â±klÃƒâ€Ã‚Â±k KatsayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"),
        Kurtosis          = T_("Kurtosis",          "BasÃƒâ€Ã‚Â±klÃƒâ€Ã‚Â±k KatsayÃƒâ€Ã‚Â±sÃƒâ€Ã‚Â±"),
        AvgDifficulty     = T_("AvgDifficulty",     "Ortalama GÃƒÆ’Ã‚Â¼ÃƒÆ’Ã‚Â§lÃƒÆ’Ã‚Â¼k (p)"),
        AvgDiscrimination = T_("AvgDiscrimination", "Ortalama Madde AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i (Corrected Item-Total)"),
        # Psychometric satiri kaldirildi
        AvgDiscriminationUL27 = T_("AvgDiscriminationUL27", "Ortalama Madde AyÃƒâ€Ã‚Â±rt EdiciliÃƒâ€Ã…Â¸i (%27lik Alt-ÃƒÆ’Ã…â€œst Grup BazlÃƒâ€Ã‚Â±)"),
        Reliability       = rel_lbl,
        MaxWeighted       = T_("MaxWeighted",       "AlÃƒâ€Ã‚Â±nabilecek Maksimum Puan"),
        WeightedExamMean  = T_("WeightedExamMean",  "SÃƒâ€Ã‚Â±nav OrtalamasÃƒâ€Ã‚Â± (katsayÃƒâ€Ã‚Â±lÃƒâ€Ã‚Â±)")
      )

      vals <- c(
        StudentsCount     = round(s$n_stu, 0),
        Median            = round(s$median, 2),
        Mode              = round(s$mode, 2),
        Mean              = round(s$mean, 2),
        SD                = round(s$sd, 2),
        Skewness          = round(s$skewness, 2),
        Kurtosis          = round(s$kurtosis, 2),
        AvgDifficulty     = round(s$avg_p, 2),
        AvgDiscrimination = round(s$avg_r, 2),
        # Psychometric satiri kaldirildi
        AvgDiscriminationUL27 = round(s$avg_ul27, 2),
        Reliability       = round(rel_val, 2),
        MaxWeighted       = round(s$max_weighted, 0),
        WeightedExamMean  = round(mean(s$weighted_total, na.rm = TRUE), 2)
      )

      stat_col  <- T_("Statistic", "Ãƒâ€Ã‚Â°statistik")
      value_col <- T_("Value",     "DeÃƒâ€Ã…Â¸er")

      test_tbl_long <- tibble::tibble(
        !!stat_col  := unname(row_labels[names(vals)]),
        !!value_col := unname(vals)
      )

      doc <- read_docx()
      doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
      doc <- body_add_par(
        doc,
        paste0(exam_title(), " - ", T_("ReportHead.Summary","Test ÃƒÆ’Ã¢â‚¬â€œzeti Ãƒâ€Ã‚Â°statistikleri")),
        style = "heading 1"
      )
      doc <- add_report_metadata(doc)

      ft_ozet <- flextable::flextable(test_tbl_long) %>%
        flextable::autofit() %>%
        flextable::width(j = 1, width = 3) %>%
        flextable::width(j = 2, width = 1.5)

      doc <- body_add_flextable(doc, ft_ozet)
      print(doc, target = file)
    },
    contentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  )
}

shinyApp(ui, server)
