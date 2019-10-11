
import java.io.BufferedWriter;  
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;  
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;  
import java.io.Writer;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.sun.org.apache.xerces.internal.impl.dv.util.Base64;

import freemarker.template.Configuration;  
import freemarker.template.Template;  
import freemarker.template.TemplateException;
import sun.misc.BASE64Encoder;  
public class DocumentHandler {  
   private Configuration configuration = null;  
   public DocumentHandler() {  
      configuration = new Configuration();  
      configuration.setDefaultEncoding("utf-8");  
   }  
  
   public void createDoc(Map dataMap) {  
      // 设置模本装置方法和路径,FreeMarker支持多种模板装载方法。可以重servlet，classpath，数据库装载，   
      // 这里我们的模板是放在/com/ybhy/word包下面   
      configuration.setClassForTemplateLoading(this.getClass(),  
            "");  
      Template t = null;  
      try {  
         // test.ftl为要装载的模板   
         t = configuration.getTemplate("forOffice.ftl");
         t.setEncoding("utf-8");  
      } catch (IOException e) {  
         e.printStackTrace();  
      }  
      // 输出文档路径及名称   
      File outFile = new File("D:/test1.doc");  
      Writer out = null;  
      try {  
         out = new BufferedWriter(new OutputStreamWriter(  
                new FileOutputStream(outFile), "utf-8"));  
      } catch (Exception e1) {  
         e1.printStackTrace();  
      }  
      try {  
         t.process(dataMap, out);  
         out.close();  
      } catch (TemplateException e) {  
         e.printStackTrace();  
      } catch (IOException e) {  
         e.printStackTrace();  
      }  
   }  
   public List<Map<String,Object>> createImg() {  
	   List<Map<String,Object>> results =new ArrayList<Map<String,Object>>();
	   InputStream in;// 读取本地文件的输入流
			File file = new File("D:/rr.png");
			try {
				in = new FileInputStream(file);
				byte[] buffer=new byte[in.available()];
				in.read(buffer);
				Map<String,Object>  dr1= new HashMap<String,Object>();
				String images=Base64.encode(buffer);
				dr1.put("IMAGE", images);
				results.add(dr1);
				results.add(dr1);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		return results;  
	   }  
   public static void main(String args[]){
	   DocumentHandler dh = new DocumentHandler();
	   Map<String,Object> dataMap1 = new HashMap<String,Object>();
	   Map<String,Object> dataMap2= new HashMap<String,Object>();
	   Map<String,Object> dataMap = new HashMap<String,Object>();
	   dataMap.put("TITLE_NAME", "张三dhjadjfkhakdjshfkadjshfkjash");  
	   dataMap.put("MUDEL_NAME", "hasdhfkashfkasdjhfsd");  
	   dataMap.put("CONTENT_NAME", "jadfllsjflsdjfldsja");
	   dataMap.put("BUG_STATUS", "第lsjdflkdjasflkjds");
	   dataMap.put("CREATE_DATE", "唱djlasjflsdjlk");
	   dataMap.put("BUG_TYPE", "东城区hhdkjshfajhdskajh北口");
	   dataMap.put("IMGLIST",dh.createImg());
	   List<Map<String,Object>> results =new ArrayList<Map<String,Object>>();
	   for(int i=0;i<3;i++) {
		   results.add(dataMap);
	   }
	   dataMap1.put("LIST", results);
	   dataMap2.put("domain", dataMap1);
	 
	   dh.createDoc(dataMap2);
	   
   }
  
}  
