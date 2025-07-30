export default {
    name: 'List send-as addresses',
    key: 'listSendAsAddresses',

    async run($) {
        const addresses = {
            data: [],
            error: null,
        };

        try {
            const response = await $.http.get('https://graph.microsoft.com/v1.0/me/sendAs', {
                headers: {
                    'Content-Type': 'application/json'
                },
                additionalProperties: {
                    skipAddingBaseUrl: true, // Ensure the base URL is not added again
                }
            });

            console.log("sendAs Data:", response.data); // Debugging line

            addresses.data = response.data.value.map((address) => ({
                value: address.emailAddress.address,
                name: `${address.displayName} (${address.emailAddress.address})`
            }));
        } catch (error) {
            addresses.error = error;
            console.error("Error fetching send-as addresses:", error); // Error logging
        }

        return addresses;
    },
};
