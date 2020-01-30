import com.aspose.words.Document;

import java.io.InputStream;
import java.net.URL;

public class DocumentLoader {

    public Document getDocument(String name) {
        try {
            InputStream inputStream = this.getClass().getClassLoader().getResourceAsStream(name);
            return new Document(inputStream);
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    public String getPath() {
        URL url = this.getClass().getResource("/");
        return url.toString();
    }
}
