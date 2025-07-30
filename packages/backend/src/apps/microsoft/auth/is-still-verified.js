const isStillVerified = async ($) => {
  try {
    const response = await $.http.get('https://graph.microsoft.com/v1.0/me', {
      additionalProperties: {
        skipAddingBaseUrl: true,
      }
    });
    return response.status === 200;
  } catch (error) {
    console.log('Error while verifying credentials', error);
    return false;
  }
};

export default isStillVerified;
