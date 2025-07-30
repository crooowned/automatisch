const setBaseUrl = ($, requestConfig) => {
  if (requestConfig.additionalProperties?.skipAddingBaseUrl)
    return requestConfig;

  requestConfig.baseURL = $.app.baseUrl;
  
  return requestConfig;
};

export default setBaseUrl;
