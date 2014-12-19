


	import java.io.File;
	import java.util.ArrayList;
	import java.util.List;


	public class ReadFiles {
		
		
		public String readFile(String filepath){
			
			String output = new String();
			int numberOfPass = 0;
			int numberOfFail = 0;
			
			File file = new File(filepath);
			File[] files = file.listFiles();
			List<File> list = new ArrayList<File>();
			
			for(File iter: files){
				
				if(iter.isDirectory()){
					readFile(iter.getAbsolutePath());
				}else{
					list.add(iter);
				}			
			}
			
			
			for(int i=0; i<list.size(); i++){
				if(list.get(i).equals("Pass")){
					numberOfPass++;
				}
				if(list.get(i).equals("Fail")){
					numberOfFail++;
				}
			}
			
			
			output = "Number of test case: "+list.size()+" Pass: "+numberOfPass+" Fail: "+numberOfFail;
			return output;
			
			
			
		}



}
