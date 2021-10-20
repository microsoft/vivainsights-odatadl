# Load required libraries
library(AzureAuth)
library(httr)
library(dplyr)
library(purrr)
library(readr)

# Enter parameters here
par_authority <- '<your-tenant>'
par_client_cred <- '<secret>'
par_client_id <- '<client-id-or-app-id>'
par_odata <- 'odata-link>'
par_outfile <- 'query_data_from_R.csv'

# Request token
token <-
  get_azure_token(
    resource = "https://workplaceanalytics.office.com",
    tenant = par_authority,
    app = par_client_id, # app ID
    password = par_client_cred, # client secret
    auth_type = "client_credentials"
  )

# Get response
r <- httr::GET(
  url = par_odata,
  httr::add_headers(
      Accept = "application/json",
      Authorization = paste("Bearer", token$credentials$access_token)
    )
  )

# Parse and clean response
# Save as multiple objects for easier debugging
jsonlite::fromJSON(r)
r_content <- httr::content(r)
r_content_value <- r_content$value

# Output file as a data frame
out_df <-
  r_content_value %>%
  purrr::map(function(x){
    
    x %>%
      purrr::flatten() %>%
      dplyr::as_tibble()
    
    }) %>%
  dplyr::bind_rows()

## Write as CSV
out_df %>%
  readr::write_csv(
    file = par_outfile
  )


