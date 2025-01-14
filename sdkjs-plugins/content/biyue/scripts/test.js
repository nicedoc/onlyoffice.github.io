function handler  {
    this.range = [];

    this.pos = []

    startElement = function(name, attributes, pos){
        if(name == "document"){            
        } else if(name == "paragraph"){
            print("Test started");
        } else if (name = "run") {

        }
        else if (name = "text") {
            print("Text started");
            str += attributes["text"];
            
            if (attributes["autonumber"] == true)
            if (attributes["underline"] == true)
            

         }
        else if (name == "drawing") {

        }

    },

    endElement = function(name){
        if(name == "test"){
            print("Test ended");
        }
    },
}