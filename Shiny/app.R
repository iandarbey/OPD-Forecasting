#
# This is a Shiny web application. You can run the application by clicking
# the 'Run App' button above.
#
# Find out more about building applications with Shiny here:
#
#    http://shiny.rstudio.com/
#

library(shiny)
library(tidyverse)
library(plotly)
library(lubridate)
load("DiverForecasts.RData")
specialtys <- unique(combined$BookSpecialty)

# Define UI for application that draws a histogram
ui <- fluidPage(

    # Application title
    titlePanel("Beaumont OPD TimeSeries"),
    # Sidebar with a slider input for number of bins 
    fluidRow(
        column(12,
            selectInput("specialty", "Pick a Specialty", choices = specialtys)
        ),

        # Show a plot of the generated distribution
        mainPanel(plotlyOutput("Plot", height = "750px", width = "1850px")
        )
    )
)

# Define server logic required to draw a histogram
server <- function(input, output) {
    load("DiverForecasts.RData")
    specialtys <- unique(combined$BookSpecialty)
    DF <- reactive({
        combined <- dplyr::filter(combined, BookSpecialty == input$specialty) %>%
            mutate(Forecast = round(Forecast,0), OvCapAvg = round(OvCapAvg,0),
                   `80% Lower` = round(`80% Lower`,0), `80% Upper` = round(`80% Upper`,0) )
    })
    output$Plot <- renderPlotly({
        print(
            ggplotly(
                ggplot(data = DF(), aes(x = MonthEnding))+
                        geom_line(aes(y = Referrals, col = 'Referrals'), size = 1.1)+
                        geom_ribbon(aes(ymin = `80% Lower`, ymax = `80% Upper`), col = "grey50", fill = "grey50", alpha = 0.3)+
                        geom_line(aes(y = Forecast, col = 'Forecast'), size = 1.1)+
                        geom_line(aes(y = OvCapAvg, col = 'OvCapAvg'), size = 1.1)+
                        theme_minimal()+
                        expand_limits(y=0)+
                        labs(title =  paste0(DF()$BookSpecialty[1],
                                             "  -  ",
                                             month(nth(DF()$MonthEnding,-7),label = TRUE),
                                             " ",
                                             year(nth(DF()$MonthEnding,-7)),
                                             " - ",
                                             "Forecast - ",
                                             round(nth(DF()$Forecast,-7),0),
                                             "  -  Capacity - ",
                                             round(nth(DF()$OvCapAvg,-7),0)),
                             x = "Month Ending",
                             color = "Measure")+
                        geom_point(data = DF()[nrow(DF())-6,],
                                   aes(x = MonthEnding,
                                       y = Forecast),
                                   fill = 'red',
                                   shape = 23,
                                   size = 5)+
                        geom_point(data = DF()[nrow(DF())-6,],
                                   aes(x = MonthEnding,
                                       y = OvCapAvg),
                                   fill = 'blue',
                                   shape = 23,
                                   size = 5)+
                        scale_color_manual(values = c("Referrals" = 'red',
                                                      "Forecast" = 'red',
                                                      "OvCapAvg" = 'blue'))+
                        theme_bw()+
                        theme_update(legend.position = "top",
                                     text = element_text(family = "serif"),
                                     plot.title = element_text(hjust = 0.5, face = "bold"),
                                     plot.subtitle = element_text(hjust = 0.5,face = "bold"),
                                     axis.title.x = element_text(face = "bold"),
                                     axis.title.y = element_text(face = "bold"))
            )
            )
        })
}

# Run the application 
shinyApp(ui = ui, server = server)
