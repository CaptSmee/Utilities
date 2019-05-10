/**
 * Used to connect to Microsoft AD and query for a users attributes using either
 * the users sAMAccountName or email.
 * Class must be run on an application server or the proper certificates must be
 * present on the machine that the class is being run.
 * Tested on Weblogic 12.2.1.2/3
 * @author DCConway
 */
public class ADUserLookup {
    
    private DirContext ctx;
    private String domainBase = "dc=,dc=,dc=,dc=";
    private SearchControls searchCtls;
    private String baseFilter = "(&((&(objectCategory=Person)(objectClass=User)))";
    private ArrayList<String> userTestInfo = new ArrayList<>();
    
    /**
     * Default Constructor for use on the Weblogic Multitenancy servers.
     * 
     * Does not create an LDAP connection.
     */
    public ADUserLookup(){

    }
    /**
     * Parameterized constructor to create a connection to ldap.
     * 
     * @param domainBase
     * @param returnAttributes
     * @param envConfig <br/><The envConfig Hashtable must contain the following entries.>
     * Context.INITIAL_CONTEXT_FACTORY, "com.sun.jndi.ldap.LdapCtxFactory"
     * Context.PROVIDER_URL, "ldap://"
     * Context.SECURITY_PROTOCOL, "ssl"
     * Context.SECURITY_PRINCIPAL, "CN=......."
     * Context.SECURITY_CREDENTIALS, "use the Tools class"
     */
    public ADUserLookup(String domainBase,String[] returnAttributes,Hashtable envConfig){
        // create ldap dir context
        this.domainBase = domainBase;
        try {
            ctx = new InitialDirContext(envConfig);
            searchCtls = new SearchControls();
            searchCtls.setSearchScope(SearchControls.SUBTREE_SCOPE);
            if(returnAttributes != null){
                searchCtls.setReturningAttributes(returnAttributes);
            }else{
                searchCtls.setReturningAttributes(null);
            }
            Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "********* Connected to AD....");
        } catch (NamingException ex) {
            Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "Exception Connecting to Active Directory", ex);
            ex.printStackTrace();
        }
    }
    
    /**
     * 
     * @param searchValue
     * @param searchBy
     * @param searchBase
     * @return
     * @throws NamingException 
     */
    public NamingEnumeration searchUser( String searchValue, String searchBy) throws NamingException {
        String filter = createQueryFilter(searchValue, searchBy);
        return this.ctx.search(this.domainBase, filter, this.searchCtls);
    }
    /**
     * Closes connection to LDAP
     */
    public void closeLdapConnection() {
        Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "Closing LDAP Connection.");
        try {
            if (ctx != null) {
                ctx.close();
            }
        } catch (NamingException e) {
            Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "Exception closing LDAP Connection", e);
        }
    }/**
     * Searches for a user in the AD by their username and returns a HashMap containing
     * a specific set of attributes. If the attribute does not exist for the user then 
     * that attribute in the map will be an empty string.
     * The attributes returned in the HashMap are:
     * - userName
     * - firstName
     * - lastName
     * - mail
     * - displayName
     * - city
     * - street
     * - state
     * - zip
     * - phone
     * - title
     * - country
     * - countryName
     * - mi
     * - company
     * - department
     * 
     * @param userName
     * @return 
     */
    public HashMap<String,String> getUserAttributesMap(String userName){
        HashMap<String,String> attributeMap = new HashMap<>();
        try {
            NamingEnumeration<SearchResult> result = searchUser(userName,"username");
            if(result.hasMore()){
                String temp;
                SearchResult rs = (SearchResult)result.next();
                Attributes attrs = rs.getAttributes();
                /**
                 * A user in the AD may or may not have a value for all of the attributes
                 * that are needed for the application. In order to prevent bad behavior 
                 * by the application if there is no value for an attribute then an
                 * empty String for that attribute is placed into the returned HashMap.
                 */
                // UserName
                try{
                    temp = attrs.get("samaccountname").toString();
                    attributeMap.put("userName", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "userName not found for user: {0}", userName);
                    attributeMap.put("userName", "");
                }
                // FirstName
                try{
                    temp = attrs.get("givenname").toString();
                    attributeMap.put("firstName", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "firstName not found for user: ", ne);
                    attributeMap.put("firstName", "");
                }
                // LastName
                try{
                    temp = attrs.get("sn").toString();
                    attributeMap.put("lastName", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "lastName not found for user", ne);
                    attributeMap.put("lastName", "");
                }
                // email
                try{
                    temp = attrs.get("mail").toString();
                    attributeMap.put("mail", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "mail not found for user", ne);
                    attributeMap.put("mail", "");
                }
                // display name
                try{
                    temp = attrs.get("cn").toString();
                    attributeMap.put("displayName", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "displayName not found for user", ne);
                    attributeMap.put("displayName", "");
                }
                // location/city
                try{
                    temp = attrs.get("l").toString();
                    attributeMap.put("city", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "city not found for user", ne);
                    attributeMap.put("city", "");
                }
                // street name
                try{
                    temp = attrs.get("StreetAddress").toString();
                    attributeMap.put("street", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "street not found for user", ne);
                    attributeMap.put("street", "");
                }
                // state
                try{
                    temp = attrs.get("st").toString();
                    attributeMap.put("state", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "state not found for user", ne);
                    attributeMap.put("state", "");
                }
                // zip code
                try{
                    temp = attrs.get("PostalCode").toString();
                    attributeMap.put("zip", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "zip not found for user", ne);
                    attributeMap.put("zip", "");
                }
                // phone number
                try{
                    temp = attrs.get("telephoneNumber").toString();
                    attributeMap.put("phone", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "phone not found for user", ne);
                    attributeMap.put("phone", "");
                }
                // title (CTR/CIV)
                try{
                    temp = attrs.get("Title").toString();
                    attributeMap.put("title", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "title not found for user", ne);
                    attributeMap.put("title", "");
                }
                // country abbreviation
                try{
                    temp = attrs.get("c").toString();
                    attributeMap.put("country", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "country not found for user", ne);
                    attributeMap.put("country", "");
                }
                // country name full
                try{
                    temp = attrs.get("co").toString();
                    attributeMap.put("countryName", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "countryName not found for user", ne);
                    attributeMap.put("countryName", "");
                }
                // middle initial
                try{
                    temp = attrs.get("Initials").toString();
                    attributeMap.put("mi", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "mi not found for user", ne);
                    attributeMap.put("mi", "");
                }
                // company - might not exist for all users
                try{
                    temp = attrs.get("Company").toString();
                    attributeMap.put("company", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "company not found for user", ne);
                    attributeMap.put("company", "");
                }
                // department - might not exist for all users
                try{
                    temp = attrs.get("Department").toString();
                    attributeMap.put("department", temp.substring(temp.indexOf(":")+1));
                }catch (Exception ne){
                    Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "department not found for user", ne);
                    attributeMap.put("department", "");
                }
            }
        } catch (NamingException ex) {
            Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "AD Query Exception", ex);
        }
        return attributeMap;
    }
    /**
     * 
     * @param userName
     * @return 
     */
    public ArrayList<String> getUserAttributes(String userName){
        try {
            NamingEnumeration<SearchResult> result = searchUser(userName,"username");
            if(result.hasMore()){
                String temp;
                SearchResult rs = (SearchResult)result.next();
                Attributes attrs = rs.getAttributes();
                temp = attrs.get("samaccountname").toString();
                userTestInfo.add("UserName: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("givenname").toString();
                userTestInfo.add("FirstName: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("sn").toString();
                userTestInfo.add("LastName: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("mail").toString();
                userTestInfo.add("Email: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("cn").toString();
                userTestInfo.add("DisplayName: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("l").toString();
                userTestInfo.add("City: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("StreetAddress").toString();
                userTestInfo.add("Street Address: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("st").toString();
                userTestInfo.add("State: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("PostalCode").toString();
                userTestInfo.add("Zip: " + temp.substring(temp.indexOf(":")+1));
                
                temp = attrs.get("telephoneNumber").toString();
                userTestInfo.add("Phone: " + temp.substring(temp.indexOf(":")+1));
               
                temp = attrs.get("Title").toString();
                userTestInfo.add("Title:: " + temp.substring(temp.indexOf(":")+1));
            }
        } catch (NamingException ex) {
            Logger.getLogger(ADUserLookup.class.getName()).log(Level.SEVERE, "AD Query Exception", ex);
        }
        return this.userTestInfo;
    }    
    private String createQueryFilter( String searchValue, String searchBy ) {
        String filter = this.baseFilter;
        if (searchBy.equals("email")) {
            filter = filter.concat("(mail=").concat(searchValue).concat("))");
        } else if (searchBy.equals("username")) {
            filter = filter.concat("(samaccountname=").concat(searchValue).concat("))");
        }
        return filter;
    }
    
    private String getDomainBase() {
       return this.domainBase;
    }
}
