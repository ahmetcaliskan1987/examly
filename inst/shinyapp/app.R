# SMART - Statistical Metrics and Reporting Tool
library(shiny)
library(dplyr); library(tidyr); library(purrr); library(stringr)
library(readr); library(readxl)
library(officer); library(flextable); library(glue)
library(magrittr)

`%||%` <- function(a, b) if (is.null(a)) b else a

d_mode <- function(x){ z <- table(x); if(length(z)==0) return(NA); as.numeric(names(z)[which.max(z)]) }
kr20 <- function(m){
  if(is.null(m)||ncol(m)<2) return(NA_real_)
  k<-ncol(m); tot<-rowSums(m,na.rm=TRUE)
  vt<-stats::var(tot,na.rm=TRUE); p<-colMeans(m,na.rm=TRUE); q<-1-p; s2<-sum(p*q,na.rm=TRUE)
  if(is.na(vt)||vt==0) return(NA_real_)
  (k/(k-1))*(1 - s2/vt)
}
pbiserial_rest <- function(item, rest){
  if(all(is.na(item)) || length(unique(na.omit(item)))<2) return(NA_real_)
  tryCatch(stats::cor(item, rest, use="pairwise.complete.obs"), error=function(e) NA_real_)
}
norm_letter <- function(x){
  v <- as.character(x) %>% str_to_upper() %>% str_trim()
  ifelse(v%in%c("A","B","C","D","E"), v, NA)
}
q_index <- function(p) 1-p
student_counts <- function(sc){
  df <- as.data.frame(sc)
  if (ncol(df)==0) return(tibble(Dogru=rep(0, nrow(df)), Yanlis=rep(0, nrow(df)), Bos=rep(0, nrow(df))))
  tibble(
    Dogru = unname(rowSums(df == 1, na.rm = TRUE)),
    Yanlis = unname(rowSums(df == 0, na.rm = TRUE)),
    Bos    = unname(rowSums(is.na(df)))
  )
}

parse_lc_raw <- function(x){
  v <- suppressWarnings(as.numeric(as.character(x)))
  v[is.na(v)] <- NA_real_
  v[v < 0 | v > 1e6] <- NA_real_
  v
}

norm_cols <- function(x){
  if (is.null(x)) return(character(0))
  if (length(x) == 1) {
    y <- unlist(strsplit(as.character(x), "[;,]+"))
    y <- trimws(y)
    y[y != ""]
  } else as.character(x)
}

get_itemexam_quant <- function(){
  q <- 0.27
  if (requireNamespace("psychometric", quietly = TRUE)) {
    res <- tryCatch(as.numeric(formals(psychometric::item.exam)$quant), error = function(e) NULL)
    if (!is.null(res) && length(res) == 1 && !is.na(res)) q <- res
  }
  return(q)
}
color_badge <- function(v, type = c("generic", "p", "r")){
  type <- match.arg(type)
  col <- "#888"
  if(!is.na(v)){
    if(type == "p"){
      col <- if(v < .40 || v > .80) "#d62728" else "#2ca02c"
    } else if(type == "r"){
      col <- if(v < .30) "#d62728" else "#2ca02c"
    } else {
      col <- if(v < .50) "#d62728" else "#2ca02c"
    }
  }
  htmltools::span(style = paste0("padding:4px 8px;border-radius:8px;background:", col, ";color:white;"),
                  sprintf("%.3f", v))
}

comment_overall_keys <- function(ap, ar){
  pkey <- if (is.na(ap)) NULL else if (ap < .4)
    "comment_overall.difficulty.low"
  else if (ap > .8)
    "comment_overall.difficulty.high"
  else
    "comment_overall.difficulty.mid"

  rkey <- if (is.na(ar)) NULL else if (ar < .2)
    "comment_overall.discrimination.low"
  else if (ar > .4)
    "comment_overall.discrimination.high"
  else
    "comment_overall.discrimination.mid"

  c(pkey, rkey)
}

difficulty_label_key <- function(p){
  if (is.na(p)) return(NA_character_)
  if (p < .40)      "labels.difficulty.hard"
  else if (p <= .80)"labels.difficulty.medium"
  else              "labels.difficulty.easy"
}

discrimination_decision_key <- function(r){
  if (is.na(r)) return(NA_character_)
  if (r < .20)      "labels.discrimination.remove"
  else if (r < .30) "labels.discrimination.consider_remove"
  else              "labels.discrimination.keep"
}

detect_id_cols <- function(cols){
  if(length(cols)==0) return(character(0))
  rx <- "(^id$|kimlik|ogrenci|öğrenci|no$|numara$|student\\s*id)"
  cols[grepl(rx, cols, ignore.case = TRUE, perl = TRUE)]
}

title_default <- "Sınav Başlığı"

ui <- navbarPage(
  title = uiOutput("app_brand"),
  windowTitle = "SMART: Statistical Metrics and Reporting Tool",
  id = "mainnav",
  header = tagList(
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
        uiOutput("lang_selector"),
        textInput("exam_title", "Sınav Başlığı", value = title_default),
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
        h3(textOutput("title_display")),
        tableOutput("dimensions_info"),
        tableOutput("sample_head")
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
        details.toggle-card > summary::after {content:'▼'; position:absolute; right:12px; top:50%; transform:translateY(-50%);}
        details.toggle-card[open] > summary::after {content:'▲';}
        details.toggle-card .toggle-body {padding:6px 14px 12px;}
        details.toggle-card .hint {margin:-6px 14px 8px 14px; color:#6b7280;}
        .banner {margin:10px 0; padding:10px 14px; background:#eef6ff; border-left:4px solid #1d4ed8; border-radius:8px; font-size:16px;}
        .details-summary-note {font-size: 12px; font-weight: 400; color: #6b7280;}
        details.toggle-card > summary { display: flex; align-items: center; gap: 8px;}
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
                                   fluidRow(
                                     column(3, uiOutput("mc_item_selector")),
                                     column(3, uiOutput("tf_item_selector")),
                                     column(3, uiOutput("bin_item_selector")),
                                     column(3, uiOutput("lc_item_selector"))
                                   )
                          )
             ),
             uiOutput("max_score_banner"),
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
             h4(textOutput("weights_header")),
             div(class="compact-wrap section-gap", uiOutput("item_weight_table")),
             br()
           )
  ),

  tabPanel(
    title = textOutput("tab_summary", container = span),
    fluidPage(
      h3(textOutput("title_display3")),
      fluidRow(
        column(12, tableOutput("test_summary_table"))
      ),
      hr(),
      fluidRow(
        column(
          12,
          h4(textOutput("summary_colour_scale")),
          tags$ul(
            tags$li(
              HTML(
                "<span style='display:inline-block;width:14px;height:14px;background:#d62728;border-radius:4px;margin-right:6px;'></span>"
              ),
              textOutput("low_label", container = span)
            ),
            tags$li(
              HTML(
                "<span style='display:inline-block;width:14px;height:14px;background:#2ca02c;border-radius:4px;margin-right:6px;'></span>"
              ),
              textOutput("high_label", container = span)
            )
          ),
          h4(textOutput("visual_cues")),
          htmlOutput("avg_p_badge"),
          br(),
          htmlOutput("avg_r_badge"),
          br(),
          textOutput("weighted_mean_text")
        )
      ),
      hr(),
      h4(textOutput("what_it_means")),
      uiOutput("overall_comment")
    )
  ),

  tabPanel(
    title = textOutput("tab_itemstats", container = span),
    fluidPage(
      h3(textOutput("title_display4")),
      tableOutput("item_stats_table"),
      hr(),
      h4(textOutput("what_it_means_items")),
      uiOutput("item_comment_ui")
    )
  ),

  tabPanel(
    title = textOutput("tab_distractor", container = span),
    fluidPage(
      h3(textOutput("title_display8")),
      uiOutput("distractor_item_picker"),
      hr(), h4(textOutput("distractor_comment_hdr")),
      verbatimTextOutput("distractor_comment"),
      hr(), h4(textOutput("distractor_table_hdr")),
      tableOutput("distractor_table")
    )
  ),

  tabPanel(
    title = textOutput("tab_student_results", container = span),
    fluidPage(
      h3(textOutput("title_display5")),
      uiOutput("student_table_html")
    )
  ),

  tabPanel(
    title = textOutput("tab_reports", container = span),
    fluidPage(
      h3(textOutput("title_display_rapor")),
      uiOutput("reports_intro_ui"),
      br(),
      uiOutput("reports_buttons_ui")
    )
  ),
  footer = tagList(
    tags$style(HTML("
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

  dict <- reactiveVal(SMART::i18n_load("tr"))
  T_   <- function(key, default = NULL) i18n_t(dict(), key, default)

  observeEvent(input$lang, {
    dict(SMART::i18n_load(input$lang))
  }, ignoreInit = FALSE)

  output$app_brand <- renderUI({
    dict()
    tags$span(T_("AppTitle", "SMART: Statistical Metrics and Reporting Tool"))
  })

  observe({
    dict()
    session$sendCustomMessage(
      "set-title",
      T_("AppTitle", "SMART: Statistical Metrics and Reporting Tool")
    )
  })

  output$tab_upload <- renderText( T_("UploadTab", "Veri Girişi") )
  output$tab_keys <- renderText( T_("KeysTab", "Madde Türleri & Anahtarlar") )
  output$set_item_types      <- renderText( T_("SetItemTypes", "Madde Tiplerini Ayarla") )
  output$set_item_types_note <- renderText( T_("SetItemTypesNote",
                                               "(Tüm maddeler 'Çoktan Seçmeli' kabul edilir. Madde tiplerini değiştirmek için tıklayın.)") )
  output$tab_summary <- renderText(T_("TestSummary", "Test Özeti"))
  output$summary_colour_scale <- renderText(T_("SummaryColourScale", "Renk Skalası"))
  output$visual_cues          <- renderText(T_("VisualCues", "Görsel Göstergeler"))
  output$what_it_means        <- renderText(T_("WhatItMeans", "Ne anlama geliyor?"))
  output$tab_itemstats      <- renderText({ T_("ItemStats", "Madde İstatistikleri") })
  output$what_it_means_items<- renderText({ T_("WhatItMeans", "Ne anlama geliyor?") })
  output$tab_distractor         <- renderText( T_("Distractor", "Çeldirici Analizi") )
  output$distractor_comment_hdr <- renderText( T_("DistractorCommentHeader", "Yorum") )
  output$distractor_table_hdr   <- renderText( T_("DistractorTableHeader", "Seçenek Dağılımları") )
  output$tab_student_results <- renderText( T_("StudentResults", "Öğrenci Sonuçları") )

  output$tab_reports <- renderText( T_("Reports", "Raporlar") )

  output$reports_intro_ui <- renderUI({
    dict()
    tagList(
      p( T_("ReportsIntro", "Aşağıdaki butonlara tıklayarak analiz sonuçlarını DOCX olarak indirebilirsiniz.") ),
      p( em( T_("ReportsNote", "(Öğrenci ve madde sayısına göre uzun sürebilir. Lütfen bekleyiniz.)") ) )
    )
  })

  output$reports_buttons_ui <- renderUI({
    dict()
    tagList(
      downloadButton("dl_rapor_ogrenci",   T_("DownloadStudentDocx",   "Öğrenci Sonuçları (DOCX)")),
      br(), br(),
      downloadButton("dl_rapor_madde",     T_("DownloadItemDocx",      "Madde İstatistikleri (DOCX)")),
      br(), br(),
      downloadButton("dl_rapor_celdirici", T_("DownloadDistractorDocx","Çeldirici Analizi (DOCX)")),
      br(), br(),
      downloadButton("dl_rapor_ozet",      T_("DownloadSummaryDocx",   "Test Özeti İstatistikleri (DOCX)"))
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
      selected = input$lang %||% "tr",
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

  output$mc_key_header <- renderText( T_("MCKeyHeader","Çoktan Seçmeli Cevap Anahtarı") )
  output$tf_key_header <- renderText( T_("TFKeyHeader","Doğru/Yanlış Maddeleri") )
  output$lc_header     <- renderText( T_("LCHeader","Uzun Cevaplı (Açık Uçlu) Maddeler") )
  output$lc_note       <- renderText( T_("LCNote","Not: Uzun cevaplı maddelerin veri setindeki değerleri öğrencinin o soruda aldığı puan olmalıdır.") )
  output$weights_header<- renderText( T_("WeightsHeader","Madde Puan Katsayıları (Maksimum)") )

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
      selected = input$sep %||% ","
    )

    updateSelectInput(
      session, "mc_choice_count",
      label = T_("MCChoiceCount", "Seçenek sayısı")
    )
  }, ignoreInit = FALSE)

  exam_title <- reactive({ req(input$exam_title); input$exam_title })

  output$title_display  <- renderText({ exam_title() })
  output$title_display_center <- renderText({ exam_title() })
  output$title_display3 <- renderText({ exam_title() })
  output$title_display4 <- renderText({ exam_title() })
  output$title_display5 <- renderText({ exam_title() })
  output$title_display8 <- renderText({ exam_title() })
  output$title_display_rapor <- renderText({ exam_title() })

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

  observeEvent(input$about_footer, {
    dict()

    pkg_name <- "SMART"
    pkg_ver  <- tryCatch(
      as.character(utils::packageDescription(pkg_name, fields = "Version")),
      error = function(e) "1.1.0"
    )
    repo_url <- getOption("SMART.repo_url", "https://github.com/your-org/SMART")

    showModal(modalDialog(
      title     = T_("AppTitle", "SMART: Statistical Metrics and Reporting Tool"),
      easyClose = TRUE,
      footer    = modalButton(T_("Close", "Kapat")),

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
        tags$li(HTML(glue::glue('{T_("Github","GitHub")}: <a href="{repo_url}" target="_blank">{repo_url}</a>'))),
        tags$li(HTML(glue::glue('{T_("Contact","İletişim")}: <a href="mailto:ahmetcaliskan1987@gmail.com">ahmetcaliskan1987@gmail.com</a>')))
      ),

      tags$p(
        em(HTML(glue::glue('© {format(Sys.Date(), "%Y")} {T_("AllRights","Tüm hakları saklıdır.")}')))
      )
    ))
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
                    T_("FileReadError", "Dosya okunamadı. CSV biçimini ve ayraç/encoding ayarlarını kontrol edin.")))
      df

    } else if (ext %in% c("xlsx", "xls")) {
      df <- tryCatch(readxl::read_excel(input$file$datapath), error = function(e) NULL)
      validate(need(!is.null(df),
                    T_("FileReadError", "Dosya okunamadı. Excel dosyasında bir sorun olabilir.")))
      df

    } else {
      validate(need(FALSE, T_("UnsupportedFile", "Desteklenmeyen dosya türü (sadece CSV/Excel).")))
    }
  })

  output$col_selectors <- renderUI({
    dict()
    df   <- raw_data()
    cols <- names(df)

    tagList(
      selectInput(
        inputId = "name_col",
        label   = T_("NameCol", "Öğrenci Adı Sütunu (isteğe bağlı)"),
        choices = setNames(
          c("", cols),
          c(paste0("(", T_("None", "Yok"), ")"), cols)
        ),
        selected = if ("Ad" %in% cols) "Ad" else ""
      ),
      selectizeInput(
        inputId = "item_cols",
        label   = T_("ItemCols", "Madde Sütunları (maddeler)"),
        choices = cols,
        multiple = TRUE,
        selected = setdiff(cols, c("Ad","İsim","Name"))
      )
    )
  })

  data_items <- reactive({
    df <- raw_data(); req(norm_cols(input$item_cols))
    sel <- setdiff(norm_cols(input$item_cols), input$name_col)
    sel <- intersect(sel, names(df))
    if(length(sel) == 0) return(df[, 0, drop = FALSE])
    as.data.frame(df[, sel, drop=FALSE])
  })

  student_names <- reactive({
    dict()
    df <- raw_data()
    n <- nrow(df)

    base <- T_("StudentWord", "Öğrenci")

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
      T_("StudentsCount", "Öğrenci (katılımcı) sayısı"),
      T_("ItemsCount",    "Soru (madde) sayısı")
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
    taken <- union(union(input$tf_items %||% character(0), input$lc_items %||% character(0)),
                   input$bin_items %||% character(0))
    base_cols <- norm_cols(input$item_cols) %||% character(0)
    if(!is.null(input$mc_items) && length(input$mc_items)>0) input$mc_items else setdiff(base_cols, taken)
  })
  tf_items_use <- reactive({ input$tf_items %||% character(0) })
  lc_items_use <- reactive({ input$lc_items %||% character(0) })
  bin_items_use <- reactive({ input$bin_items %||% character(0) })

  output$mc_item_selector <- renderUI({
    req(input$item_cols)
    sel_now <- isolate(input$mc_items)
    checkboxGroupInput("mc_items",
                       label = T_("MCItems","Çoktan Seçmeli Maddeler"),
                       choices=input$item_cols, selected=sel_now %||% mc_items_use())
  })

  output$tf_item_selector <- renderUI({
    req(input$item_cols)
    sel_now <- isolate(input$tf_items)
    checkboxGroupInput("tf_items",
                       label = T_("TFItems","Doğru/Yanlış Maddeler"),
                       choices=input$item_cols, selected=sel_now %||% tf_items_use())
  })

  output$bin_item_selector <- renderUI({
    req(input$item_cols)
    sel_now <- isolate(input$bin_items)
    checkboxGroupInput("bin_items",
                       label = T_("BINItems","1-0 Kodlu Maddeler"),
                       choices=input$item_cols, selected=sel_now %||% bin_items_use())
  })

  output$lc_item_selector <- renderUI({
    req(input$item_cols)
    sel_now <- isolate(input$lc_items)
    checkboxGroupInput("lc_items",
                       label = T_("LCItems","Uzun Cevaplı Maddeler"),
                       choices=input$item_cols, selected=sel_now %||% lc_items_use())
  })

  mc_levels <- reactive({
    ch <- as.integer(input$mc_choice_count)[1]; if(is.na(ch)) ch <- 5
    if(ch==5) c("A","B","C","D","E") else if(ch==4) c("A","B","C","D") else c("A","B","C")
  })

  output$mc_answer_table <- renderUI({
    dict();
    hdrs <- mc_items_use()
    if(length(hdrs)==0){
      return(tags$div(class="muted", T_("ThisSectionEmpty","(Bu bölümde madde yok)")))
    }

    n_txt <- as.character(length(hdrs))
    help1 <- gsub("__N__", n_txt, T_("MCKeyPasteHelp1",
                                     "Aşağıdaki __N__ adet Çoktan Seçmeli madde için cevapları SIRA GÖZETEREK yapıştırın."))
    help2 <- T_("MCKeyPasteHelp2","Giriş Sırası:")

    tagList(
      helpText(help1),
      helpText(strong(paste(help2, paste(hdrs, collapse=", ")))),
      textAreaInput(
        "mc_key_paste",
        T_("MCKeyHeader","Çoktan Seçmeli Anahtarı (Sırayla Yapıştırın)"),
        rows = 3,
        placeholder = T_("MCKeyPlaceholder","Örn: A,B,C,D,E,A,B...")
      )
    )
  })

  output$tf_answer_table <- renderUI({
    dict()
    hdrs <- tf_items_use()
    if(length(hdrs)==0){
      return(tags$div(class="muted", T_("ThisSectionNoTF","(DY olarak işaretlenen madde yok)")))
    }

    n_txt <- as.character(length(hdrs))
    help1 <- gsub("__N__", n_txt, T_("TFKeyPasteHelp1",
                                     "Aşağıdaki __N__ adet D/Y madde için cevapları SIRA GÖZETEREK yapıştırın."))
    help2 <- T_("MCKeyPasteHelp2","Giriş Sırası:")

    tagList(
      helpText(help1),
      helpText(strong(paste(help2, paste(hdrs, collapse=", ")))),
      textAreaInput(
        "tf_key_paste",
        T_("TFKeyHeader","Doğru/Yanlış Anahtarı (Sırayla Yapıştırın)"),
        rows = 2,
        placeholder = T_("TFKeyPlaceholder","Örn: D,Y,D,D...")
      )
    )
  })

  output$item_weight_table <- renderUI({
    req(norm_cols(input$item_cols))
    hdrs <- norm_cols(input$item_cols)
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
    mc <- input$mc_items %||% character(0)
    tf <- input$tf_items %||% character(0)
    lc <- input$lc_items %||% character(0)
    bn <- input$bin_items %||% character(0)
    new_tf <- setdiff(tf, mc); new_lc <- setdiff(lc, mc); new_bn <- setdiff(bn, mc)
    if(!identical(new_tf, tf)) updateCheckboxGroupInput(session, "tf_items", selected = new_tf)
    if(!identical(new_lc, lc)) updateCheckboxGroupInput(session, "lc_items", selected = new_lc)
    if(!identical(new_bn, bn)) updateCheckboxGroupInput(session, "bin_items", selected = new_bn)
  }, ignoreInit = TRUE)
  observeEvent(input$tf_items, {
    mc <- input$mc_items %||% character(0)
    tf <- input$tf_items %||% character(0)
    lc <- input$lc_items %||% character(0)
    bn <- input$bin_items %||% character(0)
    new_mc <- setdiff(mc, tf); new_lc <- setdiff(lc, tf); new_bn <- setdiff(bn, tf)
    if(!identical(new_mc, mc)) updateCheckboxGroupInput(session, "mc_items", selected = new_mc)
    if(!identical(new_lc, lc)) updateCheckboxGroupInput(session, "lc_items", selected = new_lc)
    if(!identical(new_bn, bn)) updateCheckboxGroupInput(session, "bin_items", selected = new_bn)
  }, ignoreInit = TRUE)
  observeEvent(input$bin_items, {
    mc <- input$mc_items %||% character(0)
    tf <- input$tf_items %||% character(0)
    lc <- input$lc_items %||% character(0)
    bn <- input$bin_items %||% character(0)
    new_mc <- setdiff(mc, bn); new_tf <- setdiff(tf, bn); new_lc <- setdiff(lc, bn)
    if(!identical(new_mc, mc)) updateCheckboxGroupInput(session, "mc_items", selected = new_mc)
    if(!identical(new_tf, tf)) updateCheckboxGroupInput(session, "tf_items", selected = new_tf)
    if(!identical(new_lc, lc)) updateCheckboxGroupInput(session, "lc_items", selected = new_lc)
  }, ignoreInit = TRUE)
  observeEvent(input$lc_items, {
    mc <- input$mc_items %||% character(0)
    tf <- input$tf_items %||% character(0)
    lc <- input$lc_items %||% character(0)
    bn <- input$bin_items %||% character(0)
    new_mc <- setdiff(mc, lc); new_tf <- setdiff(tf, lc); new_bn <- setdiff(bn, lc)
    if(!identical(new_mc, mc)) updateCheckboxGroupInput(session, "mc_items", selected = new_mc)
    if(!identical(new_tf, tf)) updateCheckboxGroupInput(session, "tf_items", selected = new_tf)
    if(!identical(new_bn, bn)) updateCheckboxGroupInput(session, "bin_items", selected = new_bn)
  }, ignoreInit = TRUE)

  answer_keys <- reactive({
    mc_names <- mc_items_use()
    tf_names <- tf_items_use()
    lv <- mc_levels()

    mc <- setNames(rep(NA_character_, length(mc_names)), mc_names)
    if (length(mc_names) > 0 && isTruthy(input$mc_key_paste)) {

      raw_text_mc <- input$mc_key_paste

      has_delimiter_mc <- grepl("[,;\\n\\t\\s]", raw_text_mc, perl = TRUE)

      if (has_delimiter_mc) {
        keys_pasted <- unlist(strsplit(raw_text_mc, "[,;\\n\\t\\s]+"))
      } else {
        keys_pasted <- unlist(strsplit(raw_text_mc, ""))
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

    tf <- setNames(rep(NA_character_, length(tf_names)), tf_names)
    if (length(tf_names) > 0 && isTruthy(input$tf_key_paste)) {

      raw_text_tf <- input$tf_key_paste

      has_delimiter_tf <- grepl("[,;\\n\\t\\s]", raw_text_tf, perl = TRUE)

      if (has_delimiter_tf) {
        keys_pasted_tf <- unlist(strsplit(raw_text_tf, "[,;\\n\\t\\s]+"))
      } else {
        keys_pasted_tf <- unlist(strsplit(raw_text_tf, ""))
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
    cols <- norm_cols(input$item_cols) %||% character(0)
    w <- setNames(rep(1, length(cols)), cols)
    for(nm in cols){
      v <- suppressWarnings(as.numeric(input[[paste0("w_", nm)]]))
      if(length(v) == 0 || is.na(v)) next
      if(v >= 0) w[[nm]] <- v
    }
    w
  })

  parse_tf_bin <- function(x, key){
    norm <- function(v){
      v <- as.character(v) %>% str_to_upper() %>% str_trim()
      dplyr::case_when(v%in%c("D","DOGRU","DOĞRU","TRUE","T")~"D",
                       v%in%c("Y","YANLIS","YANLIŞ","FALSE","F")~"Y",
                       TRUE~NA_character_)
    }
    as.integer(norm(x) == norm(key))
  }
  parse_mc_bin <- function(x, key){
    xN<-as.character(x) %>% str_to_upper() %>% str_trim()
    kN<-as.character(key) %>% str_to_upper() %>% str_trim()
    xN<-ifelse(xN%in%c("A","B","C","D","E"),xN,NA); kN<-ifelse(kN%in%c("A","B","C","D","E"),kN,NA)
    as.integer(xN==kN)
  }
  parse_lc_bin <- function(x){
    v <- suppressWarnings(as.numeric(as.character(x)))
    v[!(v %in% c(0,1))] <- NA
    as.integer(v)
  }
  is_scored_01 <- function(vec){
    u <- unique(na.omit(suppressWarnings(as.numeric(as.character(vec)))))
    length(u) > 0 && all(u %in% c(0,1))
  }

  scored_items <- reactive({
    df <- data_items(); keys <- answer_keys()
    if(is.null(df) || (length(norm_cols(input$item_cols) %||% character(0))==0)) return(NULL)

    all_names <- norm_cols(input$item_cols)
    sc_bin <- matrix(NA_real_, nrow=nrow(df), ncol=length(all_names), dimnames=list(NULL, all_names)) %>% as.data.frame()
    lc_raw <- matrix(NA_real_, nrow=nrow(df), ncol=length(all_names), dimnames=list(NULL, all_names)) %>% as.data.frame()

    for(nm in lc_items_use()){
      if(nm %in% names(df)){
        lc_raw[[nm]] <- parse_lc_raw(df[[nm]])
        sc_bin[[nm]] <- NA_real_
      }
    }

    for(nm in bin_items_use()){
      if(nm %in% names(df)){
        sc_bin[[nm]] <- parse_lc_bin(df[[nm]])
      }
    }

    for(nm in setdiff(all_names, union(union(mc_items_use(), tf_items_use()),
                                       union(lc_items_use(), bin_items_use())))){
      col <- df[[nm]]
      if(is_scored_01(col)){
        sc_bin[[nm]] <- parse_lc_bin(col)
      }
    }

    for(nm in names(keys$mc)){
      if(isTruthy(keys$mc[[nm]]) && nm %in% names(df)){
        sc_bin[[nm]] <- parse_mc_bin(df[[nm]], keys$mc[[nm]])
      }
    }

    for(nm in names(keys$tf)){
      if(isTruthy(keys$tf[[nm]]) && nm %in% names(df)){
        sc_bin[[nm]] <- parse_tf_bin(df[[nm]], keys$tf[[nm]])
      }
    }

    list(sc_bin = sc_bin, lc_raw = lc_raw)
  })

  output$max_score_banner <- renderUI({
    s <- test_summary()
    if (is.null(s) || is.null(s$max_weighted)) return(NULL)
    tags$div(class = "banner",
             HTML(paste0("<b>", T_("MaxScore","Alınabilecek Maksimum Puan:"), "</b> ",
                         round(s$max_weighted, 3)))
    )
  })

  compute_item_exam <- reactive({
    si <- scored_items(); if(is.null(si)) return(NULL)
    only_dicho <- si$sc_bin
    if(requireNamespace("psychometric", quietly=TRUE) && ncol(only_dicho) > 0){
      out <- tryCatch(psychometric::item.exam(only_dicho, discrim = TRUE), error = function(e) NULL)
      return(out)
    }
    NULL
  })
  itemexam_quant_reactive <- reactive({ get_itemexam_quant() })

  test_summary <- reactive({
    dict()
    si <- scored_items()
    validate(need(!is.null(si), T_("NeedItemTypesKeys", "Lütfen madde türlerini ve anahtarları girin.")))

    sc_bin <- si$sc_bin
    lc_raw <- si$lc_raw
    n_stu <- nrow(sc_bin)

    w <- item_weights()
    all_items <- colnames(sc_bin)
    w_use <- w[all_items] %||% rep(1, length(all_items))
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

    rjx <- sapply(seq_len(ncol(Y)), function(j){
      yj <- Y[[j]]
      if(all(is.na(yj))) return(NA_real_)
      pbiserial_rest(yj, weighted_total - ifelse(is.na(yj), 0, yj))
    })
    names(rjx) <- all_items

    valid_cols <- which(colSums(!is.na(Y)) > 1)
    k_eff <- length(valid_cols)
    col_vars <- if(k_eff>0) sapply(valid_cols, function(j) stats::var(Y[[j]], na.rm = TRUE)) else numeric(0)
    var_total <- stats::var(weighted_total, na.rm = TRUE)
    alpha <- if (k_eff > 1 && !is.na(var_total) && var_total > 0) (k_eff/(k_eff - 1)) * (1 - sum(col_vars, na.rm = TRUE)/var_total) else NA_real_

    only_dicho <- sc_bin[, intersect(colnames(sc_bin), union(mc_tf_bin_names, other_names)), drop = FALSE]
    kr20_val <- if(length(lc_names)==0 && all(w_use[union(mc_tf_bin_names, other_names)]==1, na.rm=TRUE)) kr20(only_dicho) else NA_real_

    q <- itemexam_quant_reactive(); k <- max(1, floor(n_stu * q))
    ord <- order(weighted_total)
    lower <- weighted_total[ord][1:k]
    upper <- weighted_total[ord][(n_stu-k+1):n_stu]

    list(
      n_stu = n_stu,
      median = stats::median(weighted_total, na.rm=TRUE),
      mode   = suppressWarnings(d_mode(weighted_total)),
      mean   = mean(weighted_total, na.rm=TRUE),
      sd     = stats::sd(weighted_total, na.rm=TRUE),
      avg_p  = mean(p_all[!is.na(p_all)], na.rm=TRUE),
      avg_r  = mean(rjx[!is.na(rjx)], na.rm=TRUE),
      alpha  = alpha,
      kr20   = kr20_val,
      lower_max = max(lower, na.rm=TRUE),
      upper_min = min(upper, na.rm=TRUE),

      total_unweighted = rowSums(sc_bin, na.rm = TRUE),
      p = p_all, rjx = rjx,
      sc_bin = sc_bin,
      lc_raw = lc_raw,
      Y = Y,
      weighted_total = weighted_total,
      max_weighted = max_weighted,
      weights = w_use,
      k_group = k, q = q
    )
  })

  output$weighted_mean_text <- renderText({
    paste0(T_("WeightedMean", "Sınav Ortalaması:"), " ", round(summary_stats$weighted_mean, 2))
  })

  output$overall_comment <- renderText({
    if (mean_p < .3) {
      ptxt <- T_("Overall_ptxt_low", "Ortalama güçlük düşük, başka bir deyişle soruların çoğu zor.")
    } else if (mean_p < .7) {
      ptxt <- T_("Overall_ptxt_mid", "Ortalama güçlük orta, başka bir deyişle sınavın zorluk seviyesi dengeli.")
    } else {
      ptxt <- T_("Overall_ptxt_high", "Ortalama güçlük yüksek, yani soruların çoğu kolay.")
    }

    if (mean_r < .2) {
      rtxt <- T_("Overall_rtxt_low", "Ortalama ayırt edicilik düşük, sorular bilenle bilmeyeni ayırmakta zayıf olabilir.")
    } else if (mean_r < .4) {
      rtxt <- T_("Overall_rtxt_mid", "Ortalama ayırt edicilik orta düzeydedir.")
    } else {
      rtxt <- T_("Overall_rtxt_high", "Ortalama ayırt edicilik yüksektir, sorular bilenle bilmeyeni iyi ayırır.")
    }

    paste(ptxt, rtxt, sep = "\n")
  })

  output$test_summary_table <- renderTable({
    s <- test_summary()

    tbl <- tibble::tibble(
      StudentsCount      = s$n_stu,
      Median             = round(s$median, 3),
      Mode               = round(s$mode, 3),
      Mean               = round(s$mean, 3),
      SD                 = round(s$sd, 3),
      AvgDifficulty      = round(s$avg_p, 3),
      AvgDiscrimination  = round(s$avg_r, 3),
      CronbachAlpha      = round(s$alpha, 3),
      KR20               = round(s$kr20, 3),
      LowerMax           = round(s$lower_max, 3),
      UpperMin           = round(s$upper_min, 3),
      MaxWeighted        = round(s$max_weighted, 3),
      WeightedExamMean   = round(mean(s$weighted_total, na.rm = TRUE), 3)
    )

    labels <- c(
      StudentsCount     = T_("StudentsCount",    "Öğrenci Sayısı"),
      Median            = T_("Median",           "Medyan"),
      Mode              = T_("Mode",             "Mod"),
      Mean              = T_("Mean",             "Aritmetik Ortalama"),
      SD                = T_("SD",               "Standart Sapma"),
      AvgDifficulty     = T_("AvgDifficulty",    "Ortalama Güçlük (p)"),
      AvgDiscrimination = T_("AvgDiscrimination","Ortalama Ayırt Edicilik (rjx)"),
      CronbachAlpha     = T_("CronbachAlpha",    "Cronbach Alfa"),
      KR20              = T_("KR20",             "KR-20"),
      LowerMax          = T_("LowerMax",         "Alt Grup Maks. Puan"),
      UpperMin          = T_("UpperMin",         "Üst Grup Min. Puan"),
      MaxWeighted       = T_("MaxWeighted",      "Alınabilecek Maksimum Puan"),
      WeightedExamMean  = T_("WeightedExamMean", "Sınav Ortalaması (katsayılı)")
    )

    names(tbl) <- unname(labels[names(tbl)])
    tbl
  }, align = "c")

  output$avg_p_badge <- renderUI({
    s <- test_summary()
    HTML(paste(T_("AvgP","Ortalama Güçlük:"), as.character(color_badge(s$avg_p, "p"))))
  })

  output$avg_r_badge <- renderUI({
    s <- test_summary()
    HTML(paste(T_("AvgR","Ortalama Ayırt Edicilik:"), as.character(color_badge(s$avg_r, "r"))))
  })

  output$weighted_mean_text <- renderText({
    s <- test_summary()
    paste0(T_("WeightedMean", "Sınav Ortalaması:"), " ",
           round(mean(s$weighted_total, na.rm = TRUE), 3))
  })

  output$overall_comment <- renderUI({
    s <- test_summary()

    keyP <- if (is.na(s$avg_p)) NULL else if (s$avg_p < .4) "Overall_ptxt_low"
    else if (s$avg_p > .8) "Overall_ptxt_high" else "Overall_ptxt_mid"
    keyR <- if (is.na(s$avg_r)) NULL else if (s$avg_r < .2) "Overall_rtxt_low"
    else if (s$avg_r > .4) "Overall_rtxt_high" else "Overall_rtxt_mid"

    ptxt <- T_(keyP, "")
    rtxt <- T_(keyR, "")

    htmltools::HTML(paste(ptxt, rtxt, sep = "<br/>"))
  })

  output$item_stats_table <- renderTable({
    dict()

    s <- test_summary()
    Y <- s$Y; p <- s$p

    item_sd  <- apply(Y, 2, stats::sd, na.rm = TRUE)
    item_var <- item_sd^2

    reliab <- sapply(seq_len(ncol(Y)), function(j){
      tryCatch(stats::cor(Y[[j]], s$weighted_total, use = "pairwise.complete.obs"),
               error = function(e) NA_real_)
    })

    diff_keys <- vapply(p,     difficulty_label_key,      NA_character_)
    disc_keys <- vapply(s$rjx, discrimination_decision_key, NA_character_)

    diff_labels <- ifelse(is.na(diff_keys), "",
                          vapply(diff_keys, function(k) T_(k, ""), character(1)))
    disc_labels <- ifelse(is.na(disc_keys), "",
                          vapply(disc_keys, function(k) T_(k, ""), character(1)))


    tbl <- tibble::tibble(
      ItemName   = names(Y),
      Difficulty = round(p, 3),
      Discrimination = round(s$rjx, 3),
      SD        = round(item_sd, 3),
      Variance  = round(item_var, 3),
      ItemReliabilityIndex = round(reliab, 3),
      DifficultyLabel = diff_labels,
      DiscriminationDecision = disc_labels
    )

    colnames(tbl) <- c(
      T_("ItemCol",                 "Soru"),
      T_("DifficultyCol",          "Güçlük"),
      T_("DiscriminationCol",      "Ayırt Edicilik"),
      T_("SD",                     "S.D."),
      T_("VarianceCol",            "Varyans"),
      T_("ItemRelIndexCol",        "Madde Güvenirlik İndeksi"),
      T_("DifficultyLabelCol",     "Zorluk"),
      T_("DiscDecisionCol",        "Madde Ayırt Edicilik Yorumu")
    )

    tbl
  }, align = "c")

  output$item_comment_ui <- renderUI({
    dict()
    l1 <- T_("ItemComment_Line1",
             "As difficulty approaches 1.00, the item is easier; as it approaches 0.00, it is harder.")
    l2 <- T_("ItemComment_Line2",
             "The discrimination coefficient reflects the item's correlation with the weighted total (corrected item-total).")
    l3 <- T_("ItemComment_Line3",
             "The Item Reliability Index is the item's correlation with the weighted total; higher values indicate positive contribution.")

    htmltools::HTML(paste(l1, l2, l3, sep = "<br/>"))
  })

  mc_raw_norm <- reactive({
    req(mc_items_use()); df <- data_items(); cols <- intersect(names(df), mc_items_use())
    if(length(cols)==0) return(NULL)
    as.data.frame(lapply(cols, function(nm) norm_letter(df[[nm]])), stringsAsFactors=FALSE) %>% setNames(cols)
  })
  compute_distractor <- function(item_name){
    s <- test_summary(); totals <- s$weighted_total
    mlv <- mc_levels(); df_mc <- mc_raw_norm(); keys <- answer_keys()$mc
    if(is.null(df_mc) || !(item_name %in% names(df_mc))) return(NULL)
    x <- df_mc[[item_name]]; key <- toupper(str_trim(keys[[item_name]]))
    n <- length(totals); q <- s$q; k <- max(1, floor(n * q)); ord <- order(totals)
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
    type_correct    <- T_("DistractorType.Correct", "Doğru")
    type_distractor <- T_("DistractorType.Distractor", "Çeldirici")

    col_option   <- T_("DistractorCol.Option",          "Seçenek")
    col_type     <- T_("DistractorCol.Type",            "Tür")
    col_cnt_all  <- T_("DistractorCol.CountAll",        "İşaretleyenlerin Sayısı")
    col_pct_all  <- T_("DistractorCol.PctAll",          "İşaretleyenlerin Yüzdesi")
    col_cnt_up   <- T_("DistractorCol.CountUpper",      "Üst Grupta İşaretleyenlerin Sayısı")
    col_pct_up   <- T_("DistractorCol.PctUpper",        "Üst Grupta İşaretleyenler %")
    col_cnt_low  <- T_("DistractorCol.CountLower",      "Alt Grupta İşaretleyenlerin Sayısı")
    col_pct_low  <- T_("DistractorCol.PctLower",        "Alt Grupta İşaretleyenler %")
    col_disc     <- T_("DistractorCol.Discrimination",  "Seçeneğin Ayırt Ediciliği")

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
    selectInput("distr_item", T_("DistractorPickItem","Madde seçin"), choices = mc_items_use())
  })

  output$distractor_comment <- renderText({
    dict(); req(input$distr_item)
    tb <- compute_distractor(input$distr_item)
    if (is.null(tb))
      return( T_("DistractorNeedsMCAndKey",
                 "Bu analiz için madde çoktan seçmeli olmalı ve cevap anahtarı girilmiş olmalıdır.") )

    type_correct    <- T_("DistractorType.Correct", "Doğru")
    type_distractor <- T_("DistractorType.Distractor", "Çeldirici")

    col_type   <- T_("DistractorCol.Type", "Tür")
    col_disc   <- T_("DistractorCol.Discrimination", "Seçeneğin Ayırt Ediciliği")
    col_option <- T_("DistractorCol.Option", "Seçenek")

    key_row <- tb %>% dplyr::filter(.data[[col_type]] == type_correct)
    ok <- nrow(key_row) > 0 && !is.na(key_row[[col_disc]][1]) && key_row[[col_disc]][1] > 0

    bads <- tb %>% dplyr::filter(.data[[col_type]] == type_distractor, .data[[col_disc]] > 0)

    msg1 <- if (ok) T_("DistractorMsg_OK",
                       "Doğru seçenek başarılı öğrenciler tarafından daha çok işaretlenmiş.")
    else     T_("DistractorMsg_NotOK",
                "Doğru seçenek başarılı öğrencilerde daha çok görünmüyor.")
    msg2 <- if (nrow(bads) == 0)
      paste0("\n", T_("DistractorMsg_OkDistractors",
                      "Çeldiriciler başarısı düşük öğrenciler tarafından daha çok seçiliyor."))
    else
      paste0("\n",
             T_("DistractorMsg_BadDistractorsPrefix",
                "Başarılı öğrencilerde görece daha çok seçilen çeldiriciler: "),
             paste0(bads[[col_option]], collapse = ", "), ".")

    paste(msg1, msg2)
  })

  output$distractor_table <- renderTable({ req(input$distr_item); compute_distractor(input$distr_item) }, align="c")

  student_item_comment <- function(sc_row, lc_row, w, mc_names, tf_names, bin_names, lc_names){
    scn <- suppressWarnings(as.numeric(sc_row)); names(scn) <- names(sc_row)
    scn[is.na(scn)] <- NA_real_
    zero_mctf <- names(scn)[names(scn) %in% union(mc_names, tf_names) & !is.na(scn) & scn == 0]
    zero_bin  <- names(scn)[names(scn) %in% bin_names                 & !is.na(scn) & scn == 0]

    lcn <- suppressWarnings(as.numeric(lc_row)); names(lcn) <- names(lc_row)
    w0  <- suppressWarnings(as.numeric(w[names(lc_row)]));   names(w0)  <- names(lc_row)
    low_lc <- names(lcn)[names(lcn) %in% lc_names & !is.na(lcn) & !is.na(w0) & (lcn < 0.5 * w0)]

    if(length(zero_mctf) == 0 && length(zero_bin) == 0 && length(low_lc) == 0) return("")
    maddeler <- paste0(sort(unique(c(zero_mctf, zero_bin, low_lc))), collapse = ", ")
    paste0(" ", maddeler, ".")
  }

  output$student_table_html <- renderUI({
    s <- test_summary()
    dict()
    if (is.null(s)) return(tags$p( T_("SummaryNotReady",
                                      "Özet hesaplanamadı. Lütfen veri ve anahtarları kontrol edin.") ))

    sc_all <- s$sc_bin
    lc_all <- s$lc_raw
    mcN <- mc_items_use(); tfN <- tf_items_use(); lcN <- lc_items_use(); bnN <- bin_items_use(); w <- s$weights

    has_mc <- length(mcN) > 0
    has_tf <- length(tfN) > 0
    has_bn <- length(bnN) > 0
    has_lc <- length(lcN) > 0

    cnt_mc <- if(has_mc) student_counts(sc_all[, intersect(colnames(sc_all), mcN), drop=FALSE]) else NULL
    cnt_tf <- if(has_tf) student_counts(sc_all[, intersect(colnames(sc_all), tfN), drop=FALSE]) else NULL
    cnt_bn <- if(has_bn) student_counts(sc_all[, intersect(colnames(sc_all), bnN), drop=FALSE]) else NULL
    lc_sum <- if(has_lc) rowSums(lc_all[, intersect(colnames(lc_all), lcN), drop=FALSE], na.rm=TRUE) else rep(0, nrow(sc_all))

    yorumlar <- purrr::map_chr(seq_len(nrow(s$sc_bin)), function(i){
      row_mctf <- s$sc_bin[i,,drop=TRUE]
      row_lc <- s$lc_raw[i,,drop=TRUE]
      student_item_comment(row_mctf, row_lc, w, mcN, tfN, bnN, lcN)
    })

    tbl_style <- "width:100%; border-collapse:collapse;"
    th_style <- "border:1px solid #e5e7eb; padding:8px; background:#f9fafb; text-align:center; font-weight:600;"
    td_style <- "border:1px solid #e5e7eb; padding:6px; text-align:center;"
    name_style <- "border:1px solid #e5e7eb; padding:6px; text-align:left;"

    dict()
    lbl_student   <- T_("StudentsColHeader", "Öğrenci")
    lbl_mc_group  <- T_("MCGroupHeader", "Çoktan Seçmeli Maddeler")
    lbl_tf_group  <- T_("TFGroupHeader", "D/Y Maddeler")
    lbl_bin_group <- T_("BINGroupHeader", "1-0 Maddeler")
    lbl_lc_group  <- T_("LCGroupHeader", "Uzun Cevaplı (Alınan Puan)")
    lbl_correct   <- T_("Correct", "Doğru")
    lbl_wrong     <- T_("Wrong",   "Yanlış")
    lbl_blank     <- T_("Blank",   "Boş")
    lbl_wscore    <- T_("WeightedScore", "Katsayılı Puan")
    lbl_weak      <- T_("StudentWeakItems", "Kazanımlara Ait Eksiklik Olabilecek Maddeler")

    header1_elements <- list(tags$th(style=th_style, rowspan=2, lbl_student))
    if (has_mc) header1_elements <- c(header1_elements, list(tags$th(style=th_style, colspan=3, lbl_mc_group)))
    if (has_tf) header1_elements <- c(header1_elements, list(tags$th(style=th_style, colspan=3, lbl_tf_group)))
    if (has_bn) header1_elements <- c(header1_elements, list(tags$th(style=th_style, colspan=3, lbl_bin_group)))
    if (has_lc) header1_elements <- c(header1_elements, list(tags$th(style=th_style, rowspan=2, lbl_lc_group)))
    header1_elements <- c(header1_elements, list(
      tags$th(style=th_style, rowspan=2, lbl_wscore),
      tags$th(style=th_style, rowspan=2, lbl_weak)
    ))
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
      row_elements <- c(row_elements, list(
        tags$td(style=td_style, round(s$weighted_total[i],3)),
        tags$td(style=paste0(td_style, " text-align:left;"), yorumlar[i])
      ))
      do.call(tags$tr, row_elements)
    })

    tags$table(style=tbl_style,
               tags$thead(header1, header2),
               tags$tbody(rows)
    )
  })

  make_student_df <- function(s, mcN, tfN, binN, lcN, names_vec){
    dict()

    sc_all <- s$sc_bin
    lc_all <- s$lc_raw

    has_mc <- length(mcN)  > 0
    has_tf <- length(tfN)  > 0
    has_bn <- length(binN) > 0
    has_lc <- length(lcN)  > 0

    col_student <- T_("StudentsColHeader", "Öğrenci")
    short_mc    <- T_("MCShort",  "ÇS")
    short_tf    <- T_("TFShort",  "D/Y")
    short_bin   <- T_("BINShort", "1-0")
    lc_sum_lbl  <- T_("LCSumShort","Uzun Cev. Toplam")

    lbl_correct <- T_("Correct", "Doğru")
    lbl_wrong   <- T_("Wrong",   "Yanlış")
    lbl_blank   <- T_("Blank",   "Boş")

    col_wscore  <- T_("WeightedScore",   "Katsayılı Puan")
    col_weak    <- T_("StudentWeakItems","Kazanımlara Ait Eksiklik Olabilecek Maddeler")

    result_list <- list()
    result_list[[col_student]] <- unname(names_vec)

    if (has_mc) {
      cnt_mc <- student_counts(sc_all[, intersect(colnames(sc_all), mcN), drop=FALSE])
      result_list[[paste(short_mc, lbl_correct)]] <- unname(cnt_mc$Dogru)
      result_list[[paste(short_mc, lbl_wrong)  ]] <- unname(cnt_mc$Yanlis)
      result_list[[paste(short_mc, lbl_blank)  ]] <- unname(cnt_mc$Bos)
    }
    if (has_tf) {
      cnt_tf <- student_counts(sc_all[, intersect(colnames(sc_all), tfN), drop=FALSE])
      result_list[[paste(short_tf, lbl_correct)]] <- unname(cnt_tf$Dogru)
      result_list[[paste(short_tf, lbl_wrong)  ]] <- unname(cnt_tf$Yanlis)
      result_list[[paste(short_tf, lbl_blank)  ]] <- unname(cnt_tf$Bos)
    }
    if (has_bn) {
      cnt_bn <- student_counts(sc_all[, intersect(colnames(sc_all), binN), drop=FALSE])
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

      short_mc   <- T_("MCShort",  "ÇS")
      short_tf   <- T_("TFShort",  "D/Y")
      short_bin  <- T_("BINShort", "1-0")
      lc_sum_lbl <- T_("LCSumShort","Uzun Cev. Toplam")

      g_mc <- T_("MCGroupHeader", "Çoktan Seçmeli Maddeler")
      g_tf <- T_("TFGroupHeader", "D/Y Maddeler")
      g_bn <- T_("BINGroupHeader","1-0 Maddeler")
      g_lc <- T_("LCGroupHeader", "Uzun Cevaplı (Alınan Puan)")

      lbl_correct <- T_("Correct","Doğru")
      lbl_wrong   <- T_("Wrong","Yanlış")
      lbl_blank   <- T_("Blank","Boş")

      has_mc <- any(df_cols %in% paste(short_mc, c(lbl_correct, lbl_wrong, lbl_blank)))
      has_tf <- any(df_cols %in% paste(short_tf, c(lbl_correct, lbl_wrong, lbl_blank)))
      has_bn <- any(df_cols %in% paste(short_bin, c(lbl_correct, lbl_wrong, lbl_blank)))
      has_lc <- lc_sum_lbl %in% df_cols

      header_values   <- list("")
      header_colwidth <- list(1)

      if (has_mc){ header_values <- c(header_values, g_mc); header_colwidth <- c(header_colwidth, 3) }
      if (has_tf){ header_values <- c(header_values, g_tf); header_colwidth <- c(header_colwidth, 3) }
      if (has_bn){ header_values <- c(header_values, g_bn); header_colwidth <- c(header_colwidth, 3) }
      if (has_lc){ header_values <- c(header_values, "") ; header_colwidth <- c(header_colwidth, 1) }

      header_values   <- c(header_values, "", "")
      header_colwidth <- c(header_colwidth, 1, 1)

      ft %>%
        add_header_row(values = unlist(header_values), colwidths = unlist(header_colwidth)) %>%
        merge_h(part = "header") %>%
        align(align = "center", part = "all") %>%
        theme_vanilla() %>%
        autofit()
    }

  make_student_docx <- function(file, title, s, mcN, tfN, binN, lcN, names_vec){
    df <- make_student_df(s, mcN, tfN, binN, lcN, names_vec)
    doc <- read_docx(); body_add_par(doc, title, style = "heading 1")
    ft <- flextable(df); ft <- add_multiheader_ft(ft)
    doc <- body_add_flextable(doc, ft)
    print(doc, target = file)
  }
  make_summary_docx <- function(file,title,s, mcN, tfN, binN, lcN, names_vec, test_tbl,item_tbl, dlist){
    doc<-read_docx(); body_add_par(doc,title,style="heading 1")
    doc<-body_add_flextable(doc,flextable(test_tbl) %>% autofit() %>% fit_to_width(6.5))
    body_add_par(doc,"\\nMadde İstatistikleri",style="heading 2")
    body_add_par(doc,"\\nÖğrenci Sonuçları",style="heading 2")
    ft_students <- flextable(make_student_df(s, mcN, tfN, binN, lcN, names_vec)) %>% add_multiheader_ft()
    doc <- body_add_flextable(doc, ft_students)

    if(length(dlist)>0){
      body_add_par(doc, "\\nÇeldirici Analizi (Seçenek Dağılımları)", style = "heading 2")
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
          stop(T_("ReportError", "Rapor oluşturulamadı. Hata Mesajı: "), "test_summary boş.")
        }
        mcN <- mc_items_use(); tfN <- tf_items_use(); binN <- bin_items_use(); lcN <- lc_items_use()
        names_vec <- student_names()
        if (is.null(names_vec) || length(names_vec) == 0) {
          stop(T_("ReportError", "Rapor oluşturulamadı. Hata Mesajı: "), "Öğrenci isimleri boş.")
        }

        df_students <- tryCatch({
          make_student_df(s, mcN, tfN, binN, lcN, names_vec)
        }, error = function(e_df) {
          stop(paste(T_("ReportError", "Rapor oluşturulamadı. Hata Mesajı: "), e_df$message))
        })

        if (is.null(df_students) || !is.data.frame(df_students) || nrow(df_students) == 0) {
          stop(T_("ReportError", "Rapor oluşturulamadı. Hata Mesajı: "), "Öğrenci veri çerçevesi boş/geçersiz.")
        }

        doc <- read_docx()
        doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
        doc <- body_add_par(
          doc,
          paste0(exam_title(), " - ", T_("ReportHead.StudentResults","Öğrenci Sonuçları")),
          style = "heading 1"
        )

        ft_students <- flextable(df_students)
        ft_students <- add_multiheader_ft(ft_students)
        ft_students <- fit_to_width(ft_students, max_width = 9.5)

        doc <- body_add_flextable(doc, ft_students)
        print(doc, target = file)

      }, error = function(e) {
        msg <- paste0(T_("ReportError", "Rapor oluşturulamadı. Hata Mesajı: "), conditionMessage(e))
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

      diff_keys <- vapply(p,     difficulty_label_key,        NA_character_)
      disc_keys <- vapply(s$rjx, discrimination_decision_key, NA_character_)

      diff_labels <- ifelse(is.na(diff_keys), "",
                            vapply(diff_keys, function(k) T_(k, ""), character(1)))
      disc_labels <- ifelse(is.na(disc_keys), "",
                            vapply(disc_keys, function(k) T_(k, ""), character(1)))

      item_tbl <- tibble::tibble(
        ItemName               = names(Y),
        Difficulty             = round(p, 3),
        Discrimination         = round(s$rjx, 3),
        SD                     = round(item_sd, 3),
        Variance               = round(item_var, 3),
        ItemReliabilityIndex   = round(reliab, 3),
        DifficultyLabel        = diff_labels,
        DiscriminationDecision = disc_labels
      )

      colnames(item_tbl) <- c(
        T_("ItemCol",            "Soru"),
        T_("DifficultyCol",      "Güçlük"),
        T_("DiscriminationCol",  "Ayırt Edicilik"),
        T_("SD",                 "S.D."),
        T_("VarianceCol",        "Varyans"),
        T_("ItemRelIndexCol",    "Madde Güvenirlik İndeksi"),
        T_("DifficultyLabelCol", "Zorluk"),
        T_("DiscDecisionCol",    "Madde Ayırt Edicilik Yorumu")
      )

      doc <- read_docx()
      doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
      doc <- body_add_par(
        doc,
        paste0(exam_title(), " - ", T_("ReportHead.ItemStats","Madde İstatistikleri")),
        style = "heading 1"
      )
      ft_items <- flextable::flextable(item_tbl) %>%
        flextable::autofit() %>%
        flextable::fit_to_width(max_width = 9.5)
      doc <- body_add_flextable(doc, ft_items)
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
        paste0(exam_title(), " - ", T_("ReportHead.Distractor","Çeldirici Analizi")),
        style = "heading 1"
      )

      if (length(mcN) == 0) {
        doc <- body_add_par(doc, T_("NoMCForDistractor","Analiz edilecek çoktan seçmeli madde yok."))
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
      dict()  # i18n
      s <- test_summary(); req(s)

      row_labels <- c(
        StudentsCount     = T_("StudentsCount",     "Öğrenci Sayısı"),
        Median            = T_("Median",            "Medyan"),
        Mode              = T_("Mode",              "Mod"),
        Mean              = T_("Mean",              "Aritmetik Ortalama"),
        SD                = T_("SD",                "Standart Sapma"),
        AvgDifficulty     = T_("AvgDifficulty",     "Ortalama Güçlük (p)"),
        AvgDiscrimination = T_("AvgDiscrimination", "Ortalama Ayırt Edicilik (rjx)"),
        CronbachAlpha     = T_("CronbachAlpha",     "Cronbach Alfa"),
        KR20              = T_("KR20",              "KR-20"),
        LowerMax          = T_("LowerMax",          "Alt Grup Maks. Puan"),
        UpperMin          = T_("UpperMin",          "Üst Grup Min. Puan"),
        MaxWeighted       = T_("MaxWeighted",       "Alınabilecek Maksimum Puan"),
        WeightedExamMean  = T_("WeightedExamMean",  "Sınav Ortalaması (katsayılı)")
      )

      vals <- c(
        StudentsCount     = round(s$n_stu, 0),
        Median            = round(s$median, 2),
        Mode              = round(s$mode, 2),
        Mean              = round(s$mean, 2),
        SD                = round(s$sd, 2),
        AvgDifficulty     = round(s$avg_p, 2),
        AvgDiscrimination = round(s$avg_r, 2),
        CronbachAlpha     = round(s$alpha, 2),
        KR20              = round(s$kr20, 2),
        LowerMax          = round(s$lower_max, 2),
        UpperMin          = round(s$upper_min, 2),
        MaxWeighted       = round(s$max_weighted, 0),
        WeightedExamMean  = round(mean(s$weighted_total, na.rm = TRUE), 2)
      )

      stat_col  <- T_("Statistic", "İstatistik")
      value_col <- T_("Value",     "Değer")

      test_tbl_long <- tibble::tibble(
        !!stat_col  := unname(row_labels[names(vals)]),
        !!value_col := unname(vals)
      )

      doc <- read_docx()
      doc <- body_set_default_section(doc, value = prop_section(page_size = page_size(orient = "landscape")))
      doc <- body_add_par(
        doc,
        paste0(exam_title(), " - ", T_("ReportHead.Summary","Test Özeti İstatistikleri")),
        style = "heading 1"
      )

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
