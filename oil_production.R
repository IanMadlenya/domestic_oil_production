# oil_production.R
# Copyright (C) 2015 Matthew Clegg.  All rights reserved.
#
# This file contains code for estimating a model of domestic oil production
# as a function of rig counts.

find_rig_count_file <- function (dir = "data") {
  # Searches for the most recently saved Baker Hughes rig count data
  files <- list.files(path=dir, pattern="^Baker.*xlsx")
  sfiles <- sort(files)
  rcfile <- sfiles[length(sfiles)]
  return(sprintf("%s/%s", dir, rcfile))
}

read_rig_counts <- function(file = find_rig_count_file(dir), dir="data", as.data.frame=FALSE) {
  # Reads the Baker Hughes rig count data from a locally stored file in XLSX format
  # Returns a zoo-indexed vector containing the total number of oil rigs
  library(zoo)
  library(readxl)
  df <- read_excel(file, sheet=2, skip=10)
  dates <- as.Date(df[,1])
  oil_rigs <- df[,ncol(df)-3]
  ok <- !is.na(dates) & !is.na(oil_rigs)  
  if (as.data.frame) {
    rig_counts <- data.frame(Date=dates[ok], Oil_rigs = oil_rigs[ok])
  } else {
    rig_counts <- zoo(oil_rigs[ok], dates[ok])
  }
  return (rig_counts)
}

download_eia_production <- function(dir="data") {
  # Downloads the most recent spreadsheet of historical crude oil production data
  # from the web site of the U.S. Energy Information Administration
  library(utils)
  source_url <- "http://ir.eia.gov/wpsr/psw01.xls"
  dest_file <- sprintf("%s/psw01-%s.xls", dir, format(Sys.Date(), "%Y%m%d"))
  download.file(source_url, dest_file)
  dest_file
}

find_eia_production_file <- function (dir = "data") {
  # Searches the local disk for the most recently downloaded spreadsheet of
  # historical crude oil production data
  files <- list.files(path=dir, pattern="^psw01.*xls")
  sfiles <- sort(files)
  eiafile <- sfiles[length(sfiles)]
  return(sprintf("%s/%s", dir, eiafile))
}

read_eia_production <- function(file = find_eia_production_file(dir), dir="data") {
  # Reads a spreadsheet file containing crude oil production data, and creates
  # a data.frame containing the following columns:
  #    Date:  The Date of the observation (typically a Friday)
  #    us_production:  This is line (1) of the Crude Oil Supply statement
  #    net_imports:    This is line (4) of the Crude Oil Supply statement
  #    stock_change:   This is line (10) of the Crude Oil Supply statement
  #    adjustment:     This is line (13) of the Crude Oil Supply statement
  #    refinery_inputs: This is line (14) of the Crude Oil Supply statement
  df2 <- read_excel(file, sheet=2, skip=2)
  dates2 <- as.Date(df2[,1])
  ok2 <- !is.na(dates2)
  crude_inventory <- df2[ok2, 3]
  ndf2 <- data.frame(Date=dates2[ok2], crude_inventory)
  ok2 <- !is.na(dates2)
  df3 <- read_excel(file, sheet=3, skip=2)
  dates3 <- as.Date(df3[,1])
  ok3 <- !is.na(dates3) 
  us_production <- df3[ok3,2]
  net_imports <- df3[ok3,5]
  stock_change <- df3[ok3,11]
  adjustment <- df3[ok3,14]
  refinery_inputs <- df3[ok3,15]
  ndf3 <- data.frame(Date=dates3[ok3], us_production, net_imports, 
                     stock_change, adjustment, refinery_inputs)
  ndf <- merge(ndf2, ndf3)
  zndf <- zoo(ndf[,2:ncol(ndf)], ndf[,1])
  zndf
}

combine_baker_hughes_and_eia <- function (bh = read_rig_counts(), eia = read_eia_production()) {
  # This function combines and harmonizes the data from the Baker Hughes rig counts
  # and the EIA crude oil supply.  Produces a data.frame containing the combined data.
  bh_dates <- as.Date(c("2012-04-05", "2012-11-21", "2013-03-28", "2013-07-03", "2013-11-27",
                        "2014-04-17", "2014-07-03", "2014-11-26", "2015-04-02", "2015-07-02"))
  eia_dates <- as.Date(c("2012-04-06", "2012-11-23", "2013-03-29", "2013-07-05", "2013-11-29",
                         "2014-04-18", "2014-07-04", "2014-11-28", "2015-04-03", "2015-07-03"))
  bhi <- index(bh)
  for (i in 1:length(bhi)) {
    if (bhi[i] %in% bh_dates) {
      bhi[i] <- eia_dates[which(bh_dates == bhi[i])]
    }
  }
  index(bh) <- bhi
  df.bhrc <- merge.zoo(bh, eia) 
  df.bhrc <- df.bhrc[(index(df.bhrc) >= index(bh)[1]) & (index(df.bhrc) >= index(eia)[1]),]
  colnames(df.bhrc)[1] <- "rig_count"
  df.bhrc
}

simulate_production <- function (bheia, 
  production_delay = 24,       # No. of weeks from drilling of well to first production
  initial_production = 0.025,    # Amount of oil produced by wells in first time period of simulation
  efficiency_growth = 1.01,     # Weekly growth factor in efficiency of drilling and production
  well_decline_rate = 0.99,    # Weekly rate at which productivity of wells declines
  starting_shale_production = 4100, # Amount of oil produced from shale in week 0 (daily barrels)
  nonshale_production = 4000,   # Daily output in barrels of stable long-term (non-shale) production
  start_date = as.Date("2014-01-03"), # First date to be used in simulation
  end_date = as.Date("2015-10-30"), # End date
  window = TRUE                # If TRUE, the resulting data.frame only contains
                                # observations between start_date and end_date
) {
  # On input, bheia is a data.frame of rig counts and EIA production data, as produced
  # by the function combine_baker_hughes_and_eia().  Simulates the actual level of
  # production that would be achieved in the period from start_date to end_date using
  # the parameters production_delay, initial_production, efficiency_growth, well_decline_rate
  # and nonshale_production.  The production in week k is given by the equations:
  #    new shale production = (# rigs in week k minus production_delay) * 
  #       (initial_production * efficiency_growth^(k-1))
  #    total shale production = new shale production + 
  #       well_decline_rate * (total shale production for week k-1)
  #    total production = total shale production + nonshale_production
  # Augments the data.frame bheia with the following columns:
  #    new_shale   = new production due to shale oil in this week
  #    total_shale = total production due to shale oil in this week
  #    total_production = total overall production in this week
  
  bheia$new_shale <- NA
  bheia$total_shale <- NA
  bheia$total_production <- NA
  
  i0 <- -1
  for (i in 1:nrow(bheia)) {
    if ((start_date <= index(bheia)[i]) && (index(bheia)[i] <= end_date)) {
      if ((i0 < 0) && (index(bheia)[i] >= start_date)) {
        i0 <- i
      }
      bheia$new_shale[i] <- as.numeric(bheia$rig_count[i-production_delay]) * initial_production * 
        efficiency_growth ^ (i - i0)
      if (i0 == i) {
        bheia$total_shale[i] <- starting_shale_production * well_decline_rate + bheia$new_shale[i]
      } else {
        bheia$total_shale[i] <- as.numeric(bheia$total_shale[i-1]) * well_decline_rate + 
          bheia$new_shale[i]
      }
      bheia$total_production[i] <- bheia$total_shale[i] + nonshale_production
    }
  }
  
  if (window) {
    bheia <- window(bheia, start=start_date, end=end_date)
  }
  bheia
}

evaluate_production_model <- function (
  bheia,
  production_delay = 24,       # No. of weeks from drilling of well to first production
  initial_production = 0.025,    # Amount of oil produced by wells in first time period of simulation
  efficiency_growth = 1.01,     # Weekly growth factor in efficiency of drilling and production
  well_decline_rate = 0.99,    # Weekly rate at which productivity of wells declines
  starting_shale_production = 3770, # Amount of oil produced from shale in week 0 (daily barrels)
  nonshale_production = 4350,   # Daily output in barrels of stable long-term (non-shale) production
  start_date = as.Date("2014-01-03"), # First date to be used in simulation
  end_date = as.Date("2015-10-30") # End date
) {
  # Calculates the mean-squared error of the production model
  # with the parameters given above.
  bheiap <- simulate_production(bheia,
                                  production_delay = production_delay,
                                  initial_production = initial_production,
                                  efficiency_growth = efficiency_growth,
                                  well_decline_rate = well_decline_rate,
                                  starting_shale_production = starting_shale_production,
                                  nonshale_production = nonshale_production,
                                  start_date = start_date,
                                  end_date = end_date)
  msqe <- mean((bheiap$us_production - bheiap$total_production)^2)
  return (msqe)
}  

evaluate_production_model_quickly <- function (
  bheia,
  production_delay = 24,       # No. of weeks from drilling of well to first production
  initial_production = 0.025,    # Amount of oil produced by wells in first time period of simulation
  efficiency_growth = 1.01,     # Weekly growth factor in efficiency of drilling and production
  well_decline_rate = 0.99,    # Weekly rate at which productivity of wells declines
  starting_shale_production = 3770, # Amount of oil produced from shale in week 0 (daily barrels)
  nonshale_production = 4350,   # Daily output in barrels of stable long-term (non-shale) production
  start_date = as.Date("2014-01-03"), # First date to be used in simulation
  end_date = as.Date("2015-10-30") # End date
) {
  # Calculates the mean-squared error of the production model
  # with the parameters given above.  This should produce the same
  # result as evaluate_production_model().

  w <- which((start_date <= index(bheia)) & (index(bheia) <= end_date))
  i0 <- w[1]
  iN <- w[length(w)]
  
  rig_count <- as.numeric(bheia$rig_count)[(i0-production_delay):(iN-production_delay)]
  new_shale <- rig_count * initial_production * efficiency_growth ^ (0:(length(rig_count)-1))
  us_production <- as.numeric(bheia$us_production[i0:iN])                                               
  ssqe <- 0
  total_shale <- starting_shale_production
  for (i in 1:(iN-i0+1)) {
      total_shale <- total_shale * well_decline_rate + new_shale[i]
      total_production <- total_shale + nonshale_production
      ssqe <- ssqe + (total_production - us_production[i])^2
  }  
  ssqe / (iN - i0 + 1)
}  

find_optimal_production_model <- function 
  (bheia, 
   production_delay=24,           # Number of weeks production delay that will be assumed
   conventional_oil_ratio = 0.5,  # Proportion of initial production to be attributed to 
                                  # conventional sources.
   start_date = as.Date("2014-01-03"), # First date to be used in simulation
   end_date = as.Date("2015-10-30"), # End date
   grid_size = 4                  # Size of one dimension in grid of starting points.
                                  # Total number of starting points = grid_size ^ 3
) {
  # For given parameters of production_delay, starting_shale_production and nonshale_production,
  # finds the optimal parameters for initial_production, efficiency_growth and well_decline_rate
  library("compiler")
  enableJIT(3)
  
  prior_date_indexes <- which(index(bheia) < start_date)
  last_prior <- prior_date_indexes[length(prior_date_indexes)]
  nonshale_production <- as.numeric(bheia$us_production[last_prior] * conventional_oil_ratio)
  starting_shale_production <- as.numeric(bheia$us_production[last_prior] * (1 - conventional_oil_ratio))
  
  lower <- c(initial_production = 0.001, efficiency_growth = 1.0, well_decline_rate=0.956)
  upper <- c(initial_production = 1.0, efficiency_growth = 1.022, well_decline_rate = 1.0)
  
  grid_lower <- c(initial_production = 0.01, efficiency_growth = 1.0, well_decline_rate = 0.96)
  grid_upper <- c(initial_production = 0.20, efficiency_growht = 1.02, well_decline_rate = 1.0)
  
  midpoints <- function(a, b, n=1) { (1:n) * (b - a) / (n+1) + a}
  gridpoints <- function(a, b) midpoints(a, b, grid_size)
  igrid <- gridpoints(grid_lower[1], grid_upper[1])
  egrid <- gridpoints(grid_lower[2], grid_upper[2])
  wgrid <- gridpoints(grid_lower[3], grid_upper[3])
  grid <- expand.grid(igrid, egrid, wgrid)
  colnames(grid) <- c("initial_production", "efficiency_growth", "well_decline_rate")
  
#  if (random.start) {
#    p0 = runif(3, lower, upper)
#  } else {
#    p0 = c(initial_production = 0.05, efficiency_growth = 1.0034, well_decline_rate=0.9853)
#  }
  
  obj <- function(p) {
    evaluate_production_model_quickly(bheia,
                                      production_delay = production_delay,
                                      initial_production = p[1],
                                      efficiency_growth = p[2],
                                      well_decline_rate = p[3],
                                      starting_shale_production = starting_shale_production,
                                      nonshale_production = nonshale_production,
                                      start_date = start_date,
                                      end_date = end_date)    
  }

  best <- NA
  for (i in 1:nrow(grid)) {
    p0 <- as.numeric(grid[i,])
    res <- optim(p0, obj, method="L-BFGS-B", lower=lower, upper=upper)
    if ((i == 1) || (res$value < best$value)) best <- res 
  }

#  print(grid)

  best$par <- c(best$par, c(starting_shale_production=starting_shale_production, nonshale_production=nonshale_production))
    
  best
}

find_initial_production <- function 
(bheia, 
 production_delay=24,
 efficiency_growth = exp(log(2)/52),
 well_decline_rate = exp(log(0.5)/52),
 conventional_oil_ratio = 0.5,  # Proportion of initial production to be attributed to 
                                # conventional sources.
 start_date = as.Date("2014-01-03"), # First date to be used in simulation
 end_date = as.Date("2015-10-30") # End date
) {
  # For given parameters of production_delay, efficienc_growth, well_decline_rate,
  # starting_shale_production and nonshale_production,
  # finds the optimal parameters for initial_production

  prior_date_indexes <- which(index(bheia) < start_date)
  last_prior <- prior_date_indexes[length(prior_date_indexes)]
  nonshale_production <- as.numeric(bheia$us_production[last_prior] * conventional_oil_ratio)
  starting_shale_production <- as.numeric(bheia$us_production[last_prior] * (1 - conventional_oil_ratio))
  
  obj <- function(p) {
    evaluate_production_model_quickly(bheia,
                                      production_delay = production_delay,
                                      initial_production = p,
                                      efficiency_growth = efficiency_growth,
                                      well_decline_rate = well_decline_rate,
                                      starting_shale_production = starting_shale_production,
                                      nonshale_production = nonshale_production,
                                      start_date = start_date,
                                      end_date = end_date)    
  }
  
  res <- optimize(obj, interval=c(1e-4,1e4))
  res$par <- c(production_delay = production_delay,
              initial_production=res$minimum,
              efficiency_growth = efficiency_growth,
              well_decline_rate = well_decline_rate,
              starting_shale_production = starting_shale_production,
              nonshale_production = nonshale_production
  )
#  res$par <- c(res$par, c(starting_shale_production=starting_shale_production, nonshale_production=nonshale_production))
  res
}

find_production_parameters_by_week <- function (
  bheia, 
  production_delays=1:36, 
  conventional_oil_ratio = 0.5,  # Proportion of initial production to be attributed to 
                                 # conventional sources.
  start_date = as.Date("2014-01-03"), # First date to be used in simulation
  end_date = as.Date("2015-10-30"), # End date
  grid_size = 4
) {
  # For each possible production delay, calculates the parameters of the best
  # fitting model.  Creates a data.frame containing the results.
  do_week <- function (i) {
    model <- find_optimal_production_model( 
       bheia, 
       production_delay=i,
       conventional_oil_ratio = conventional_oil_ratio,
       start_date = start_date, # First date to be used in simulation
       end_date = end_date, # End date
       grid_size = grid_size
      ) 
    model.df <- data.frame(production_delay = i,
                           start_date = start_date,
                           end_date = end_date,
                           initial_production = model$par[1],
                           efficiency_growth = model$par[2],
                           well_decline_rate = model$par[3],
                           starting_shale_production = model$par[4],
                           nonshale_production = model$par[5],
                           efficiency_ann = model$par[2]^52,
                           well_decline_ann = model$par[3]^52,
                           msqerr = model$value)
    rownames(model.df) <- NULL
    model.df
  }
  
  ppbw <- do.call("rbind", lapply(production_delays, do_week))
}

extend_eia_df <- function(eiabh, new_end_date = as.Date("2016-06-30"), rig_count=572) {
  # Adds NA rows to the data.frame eiabh
  nd <- index(eiabh)[length(index(eiabh))] + 7
  d <- c()
  while (nd <= new_end_date) {
    d[length(d)+1] <- nd
    nd <- nd + 7
  }
  if (length(d) == 0) return(eiabh)
  
  ndf <- data.frame(rig_count=rep(rig_count, length(d)), 
                    crude_inventory=NA,
                    us_production=NA,
                    net_imports=NA,
                    stock_change=NA,
                    adjustment=NA,
                    refinery_inputs=NA)
  
  zndf <- zoo(ndf, d)
  rbind(eiabh, zndf)
}

forecast_production <- function (eiabh, params, start_date=as.Date("2014-01-03"), forecast_date=as.Date("2016-06-30"), rig_count=572) {
  edf <- extend_eia_df(eiabh, new_end_date=forecast_date, rig_count=rig_count)
  edfs <- simulate_production(
    edf, 
    production_delay = params$production_delay, 
    initial_production = params$initial_production,
    efficiency_growth = params$efficiency_growth,
    well_decline_rate = params$well_decline_rate,
    starting_shale_production = params$starting_shale_production,
    nonshale_production = params$nonshale_production,
    start_date = start_date,
    end_date = forecast_date
  )
  edfs  
}

plot_forecast <- function (eiabh, params, start_date=as.Date("2014-01-03"), forecast_date=as.Date("2016-06-30"), rig_count=572) {
  library(ggplot2)
  edf <- extend_eia_df(eiabh, new_end_date=forecast_date, rig_count=rig_count)
  edfs <- simulate_production(
    edf, 
    production_delay = params$production_delay, 
    initial_production = params$initial_production,
    efficiency_growth = params$efficiency_growth,
    well_decline_rate = params$well_decline_rate,
    starting_shale_production = params$starting_shale_production,
    nonshale_production = params$nonshale_production,
    start_date = start_date,
    end_date = forecast_date
    )

  df.a <- data.frame(Date=index(edfs), Series="Actual", Value=edfs$us_production)
  df.b <- data.frame(Date=index(edfs), Series="Model", Value=edfs$total_production)
  df <- rbind(df.a, df.b)
  p <- ggplot(data=df, aes(x=Date, y=Value, colour=Series)) + geom_line(size=1)
  p
}

multiplot_forecast <- function (eiabh, params, start_date=as.Date("2014-01-03"), forecast_date=as.Date("2016-06-30"), rig_count=572) {
  library(ggplot2)
  edf <- extend_eia_df(eiabh, new_end_date=forecast_date, rig_count=rig_count)  
  
  df.a <- NULL
  for (i in 1:nrow(params)) {
    edfs <- simulate_production(
      edf, 
      production_delay = params[i,"production_delay"], 
      initial_production = params[i, "initial_production"],
      efficiency_growth = params[i, "efficiency_growth"],
      well_decline_rate = params[i, "well_decline_rate"],
      starting_shale_production = params[i, "starting_shale_production"],
      nonshale_production = params[i, "nonshale_production"],
      start_date=start_date,
      end_date = forecast_date
    )
    
    if (is.null(df.a)) {
      df.a <- data.frame(Date=index(edfs), Series="Actual", Value=edfs$us_production)      
    }
    
    series_name = sprintf("%d weeks", params[i, "production_delay"])
    df.b <- data.frame(Date=index(edfs), Series=series_name, Value=edfs$total_production)
    df.a <- rbind(df.a, df.b)
  }
  p <- ggplot(data=df.a, aes(x=Date, y=Value, colour=Series)) + geom_line()
  p
}

write_forecast <- function (fc, filename) {
  dates <- zoo(format(index(fc), "%m/%d/%Y"), index(fc))
  df <- cbind(Date=dates, fc)
  write.csv(df, filename)
}

find_production_minimum <- function (eiabh, params, start_date=as.Date("2014-01-03"), rig_count=572) {
  forecast_date <- start_date + 10 * 365
  end_date <- index(eiabh)[nrow(eiabh)]
  edf <- extend_eia_df(eiabh, new_end_date=forecast_date, rig_count=rig_count)
  edfs <- simulate_production(
    edf, 
    production_delay = params$production_delay, 
    initial_production = params$initial_production,
    efficiency_growth = params$efficiency_growth,
    well_decline_rate = params$well_decline_rate,
    starting_shale_production = params$starting_shale_production,
    nonshale_production = params$nonshale_production,
    start_date = start_date,
    end_date = forecast_date
  )
  edfs.later <- edfs[index(edfs) > end_date,]
  min_prod <- which.min(as.numeric(edfs.later$total_production))
  edfs.later[min_prod,]  
}

plot_adhoc_model <- function 
(bheia, 
 production_delay=24,
 efficiency_growth = exp(log(2)/52),
 well_decline_rate = exp(log(0.5)/52),
 conventional_oil_ratio = 0.5,  # Proportion of initial production to be attributed to 
 # conventional sources.
 start_date = as.Date("2014-01-03"), # First date to be used in simulation
 end_date = as.Date("2015-11-15") # End date
) {
  library(ggplot2)
  m <- find_initial_production(
    bheia,
    production_delay = production_delay,
    efficiency_growth = efficiency_growth,
    well_decline_rate = well_decline_rate,
    conventional_oil_ratio = conventional_oil_ratio,
    start_date = start_date,
    end_date = end_date)
  p <- plot_forecast(bheia, as.data.frame(t(m$par)), forecast_date=end_date)
  p <- p + ggtitle("Actual Production versus Conventional Wisdom") +
    ylab("Domestic Oil Production\n(Thousands of Barrels Per Day)") + xlab("")
  p
}

plot_best_model <- function (bheia, bheia.models) {
  library(ggplot2)
  ibest <- which.min(bheia.models$msqerr)
  p <- plot_forecast(bheia, bheia.models[ibest,], forecast_date=bheia.models$end_date[ibest])
  p <- p + ggtitle("Comparison of Actual Production to Best Fitting Model") +
    ylab("Domestic Oil Production\n(Thousands of Barrels Per Day)") + xlab("")
  p
}

plot_top_models <- function (bheia, bheia.models) {
  library(ggplot2)
  p <- multiplot_forecast(bheia, bheia.models, forecast_date=as.Date("2016-06-30"))
  p <- p + ggtitle("Forecasts of Future Production") +
    ylab("Domestic Oil Production\n(Thousands of Barrels Per Day)") + xlab("")
  p
  
}

plot_by_year <- function (DF, colname, start_date=as.Date("2011-01-01"), 
    end_date=as.Date("2015-12-31"), mean_window=4) {
    library(ggplot2)
    library(scales)
    DF$Value <- rollapply(DF[,colname], mean_window, mean, align="right", fill=NA)
    DF <- DF[index(DF) >= start_date & index(DF) <= end_date,]
    Dates <- index(DF)
    DF <- as.data.frame(DF)
    DF$Date <- Dates
    DF$Year <- format(DF$Date, "%Y")
#    DF$Jdate <- as.numeric(format(DF$Date, "%j"))
    Jdate <- as.Date(sapply(DF$Date, function(d)
        as.Date(sprintf("2015%s", format(d, "%m%d")), "%Y%m%d")
    ))
    DF$Jdate <- Jdate
    p <- ggplot(data=DF, aes(x=Jdate, y=Value, colour=Year)) + geom_line(size=1) +
        scale_x_date(labels = date_format("%b"))
    p
}

plot_net_imports_by_year <- function (bheia,
    start_date=as.Date("2011-01-01"), 
    end_date=as.Date("2015-12-31"), mean_window=4) {
    p <- plot_by_year(bheia, "net_imports", 
        start_date = start_date,
        end_date = end_date,
        mean_window = mean_window)
    p + ggtitle("Net Imports by Year") + ylab("Thousands of Barrels per Day") +
        xlab("")
}

plot_refinery_inputs_by_year <- function (bheia,
    start_date=as.Date("2011-01-01"), 
    end_date=as.Date("2015-12-31"), mean_window=4) {
    p <- plot_by_year(bheia, "refinery_inputs", 
        start_date = start_date,
        end_date = end_date,
        mean_window = mean_window)
    p + ggtitle("Refinery Inputs by Year") + ylab("Thousands of Barrels per Day") +
        xlab("")
}

plot_eia_adjustment <- function (eia,
    start_date=as.Date("2011-01-01"), 
    end_date=as.Date("2015-12-31")
) {
    eia <- eia[index(eia) >= start_date & index(eia) <= end_date,]
    eiadf <- as.data.frame(eia)
    eiadf$Date <- index(eia)
    p <- ggplot(data=eiadf, aes(x=Date, y=adjustment)) + geom_line(size=1) +
        ggtitle("Adjustment (Unaccounted for Crude Oil)") +
            xlab("") + ylab("Thousand of Barrels per Day")
    
        
    p
}

summarize_net_imports_by_year <- function(DF, ...) {
    summarize_eia_data_by_year(eia, "net_imports", ...)
}


summarize_refinery_inputs_by_year <- function(DF, ...) {
    summarize_eia_data_by_year(eia, "refinery_inputs", ...)
}

summarize_eia_data_by_year <- function(DF, colname, 
    start_date=as.Date("2010-11-01"), nyear=5) {
    start_year <- as.numeric(format(start_date, "%Y"))
    start_md <- format(start_date, "%m%d")
    cutoff_dates <- as.Date(sapply((0:nyear) + start_year, 
        function(y) as.Date(sprintf("%d%s", y, start_md), "%Y%m%d")))
#    cutoff_dates <- (0:nyear)*365 + start_date
    DF <- DF[index(DF) >= cutoff_dates[1] & index(DF) <= cutoff_dates[length(cutoff_dates)],]
    Dates <- index(DF)
    DF <- as.data.frame(DF)
    DF$Date <- Dates
    DF$Year <- format(DF$Date, "%Y")
    DF$year_index <- findInterval(DF$Date, cutoff_dates)
    DF$value <- DF[,colname]
    DF$prior_year_value <- c(rep(NA, 52), DF$value[1:(nrow(DF) - 52)])

    mean_str <- sprintf("mean(%s)", colname)
    cat(sprintf("%-25s %6s %20s %10s %10s %10s\n", "Period", "Nobs", mean_str, "Y/Y Growth", 
      "LM Growth", "LM stderr"))
     
    date_format = "%m/%d/%y"
    for (i in 1:nyear) {
        period_str <- sprintf("%s - %s", 
            format(cutoff_dates[i], date_format), format(cutoff_dates[i+1]-1, date_format))
        dfp <- DF[DF$year_index == i,]
        meanval <- mean(dfp$value, na.rm=TRUE)
        cat(sprintf("%-25s %6d %20.0f ", period_str, nrow(dfp), meanval))
        if (i > 1) {
            yoy_growth <- (meanval / prior_meanval - 1) * 100.0
            vlm <- summary(lm(value ~ prior_year_value + 0, dfp))
            lm_growth <- (coef(vlm)[1] - 1) * 100.0
            lm_stderr <- coef(vlm)[2] * 100.0
            cat(sprintf("%10.2f %10.2f %10.2f\n", yoy_growth, lm_growth, lm_stderr))
        } else {
            cat("\n")
        }
        prior_meanval <- meanval
    }
    
    mean1 <- mean(DF$value[DF$year_index == 1])
    meanN <- mean(DF$value[DF$year_index == nyear])
    cag <- 100.0 * (exp(log(meanN/mean1) / (nyear-1)) - 1.0)
    tlm <- summary(lm(value ~ prior_year_value + 0, DF))
    tlm_growth <- (coef(tlm)[1] - 1) * 100.0
    tlm_stderr <- coef(tlm)[2] * 100.0
    
    cat(sprintf("%-25s %6d %20.0f %10.2f %10.2f %10.2f\n",
        "Total", nrow(DF), mean(DF$value, na.rm=TRUE), cag, tlm_growth, tlm_stderr))        
}

plot_forecasts <- function (df1, df2, colname, 
    forecast_start = as.Date("2015-11-20"),
    start_date=as.Date("2015-01-01"), 
    window_size=1) {
    df1 <- as.data.frame(df1)
    df2 <- as.data.frame(df2)
    df1$date <- as.Date(df1$date)
    df2$date <- as.Date(df2$date)
    df1.a <- df1[df1$date <= forecast_start,c("date",colname)]
    if (window_size > 1) {
        df1.a[,2] <- rollapply(df1.a[,2],4,mean,align="right",fill=NA)
    }
    df1.a <- df1.a[df1.a$date >= start_date,]
    df1.b <- df1[df1$date >= forecast_start, c("date", colname)]
    df2.b <- df2[df2$date >= forecast_start, c("date", colname)]
    df1.a$Series = "Historical"
    df1.b$Series = "Conservative"
    df2.b$Series = "Baseline"
    df <- rbind(df1.a, df1.b, df2.b)
    colnames(df) <- c("Date", "Value", "Series")
    options(scipen=3)
    p <- ggplot(data=df, aes(x=Date, y=Value, colour=Series)) + geom_line(size=1) +
        scale_x_date(labels = date_format("%b %y")) + xlab("") +
        ylab("Thousands of Barrels Per Day")
    p
}

plot_production_forecasts <- function (df1, df2, forecast_start = as.Date("2015-11-20")) {
    p <- plot_forecasts(df1, df2, "production", forecast_start) +
        ggtitle("Forecast Crude Oil Production")
    p
}

plot_net_imports_forecasts <- function (df1, df2, forecast_start = as.Date("2015-11-20")) {
    p <- plot_forecasts(df1, df2, "net_imports", forecast_start, window_size=4) +
        ggtitle("Forecast Net Imports of Crude Oil")
    p
}
  
plot_adjustment_forecasts <- function (df1, df2, forecast_start = as.Date("2015-11-20")) {
    p <- plot_forecasts(df1, df2, "adjustment", forecast_start) +
        ggtitle("Forecast Adjustment (Unaccounted for Crude Oil)")
    p
}

plot_refinery_inputs_forecasts <- function (df1, df2, forecast_start = as.Date("2015-11-20")) {
    p <- plot_forecasts(df1, df2, "refinery_inputs", forecast_start) +
        ggtitle("Forecast Refinery Inputs")
    p
}

plot_crude_stocks_forecasts <- function (df1, df2, forecast_start = as.Date("2015-11-20")) {
    p <- plot_forecasts(df1, df2, "crude_stocks", forecast_start) +
        ggtitle("Forecast Crude Stocks")
    p
}
  
plot_and_save_all_forecasts <- function (forecast_source = "forecasts_20151125.xlsx",
    dir = ".") {
    fc1 <- read_excel("forecasts_20151125.xlsx",sheet=1)
    fc2 <- read_excel("forecasts_20151125.xlsx",sheet=2)
    
    png(sprintf("%s/%s", dir, "../production_forecasts.png"),height=240)
    print(plot_production_forecasts(fc1, fc2))
    dev.off()
    
    png(sprintf("%s/%s", dir, "../net_imports_forecast.png"), height=240)
    print(plot_net_imports_forecasts(fc1, fc2))
    dev.off()
    
    png(sprintf("%s/%s", dir, "../adjustment_forecast.png"), height=240)
    print(plot_adjustment_forecasts(fc1, fc2))
    dev.off()
    
    png(sprintf("%s/%s", dir, "../refinery_inputs_forecasts.png"), height=240)
    print(plot_refinery_inputs_forecasts(fc1, fc2))
    dev.off()
    
    png(sprintf("%s/%s", dir, "../crude_stocks_forecasts.png"), height=240)
    print(plot_crude_stocks_forecasts(fc1, fc2))
    dev.off()
}
