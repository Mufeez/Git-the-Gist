
 Office.initialize = function(reason){ };

describe("Compose Test for office api", function() {
beforeAll(function(done) { setTimeout(function(){ done();  },1000) });

    it("hostname :office Api", function() {
        
        var hostName_outlook = Office.context.mailbox.diagnostics.hostName;
        expect(hostName_outlook).toEqual("Outlook");
    });

    it("itemtype(like message):office Api", function() {
        
        
        var itemtype_outlook = Office.context.mailbox.item.itemType;
       expect(itemtype_outlook).toEqual("message");
     
    });

    it("Outlook theme():office Api", function() {
        
        
//var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
//var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor;
var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;
console.log("Body:(" + "bodyBackgroundColor" + "," +"bodyForegroundColor" + "), Control:(" + controlBackgroundColor + "," + controlForegroundColor + ")");
     
    });
})


//console.log("Display language: " + Office.context.displayLanguage);



