    
	import java.io.File;  
    import java.io.FileInputStream;  
    import java.util.Iterator;
    import java.util.*;
    import java.io.*;
    //Jar files
    import org.apache.poi.ss.usermodel.Cell;  
    import org.apache.poi.ss.usermodel.Row;  
    import org.apache.poi.xssf.usermodel.XSSFSheet;  
    import org.apache.poi.xssf.usermodel.XSSFWorkbook;  
    
    
    public class MyProject  
    {  
    	
    	static class SubjectDetails //Marking Scheme details
    	{
    		String grade[];
    		double gradePoint[]; 
    		double minMarks,credits;
    		
    		public SubjectDetails()
    		{
    			grade=new String[101];
    			Arrays.fill(grade, "F");
    			gradePoint=new double[101];
    			minMarks=0;
    			credits=0;
    		}
    	}
    	
    	
    	static class StudentDetails //Students details
    	{
    		int slno;
    		double marks,credit,gradeFP;	
    		String name,sub,eligibility,grade;
    		
    		public StudentDetails(int slno,double marks,double credit,double gradeFP,String name,String eligibility,String grade,String sub)
    		{	
    			this.credit=credit;
    			this.slno=slno;
    			this.eligibility=eligibility;
    			this.grade=grade;
    			this.gradeFP=gradeFP;
    			this.marks=marks;
    			this.name=name;
    			this.sub=sub;
    		}	
    	}
    	
    	
    	static void checker(String arr[],double minMarks,String name) // checker for errors in Ranges
    	{
    		int i =100,start=-1,end=-1;
    		
    		for(i =100;i>=minMarks;i--)
    		{
    			if(arr[i].equals("F"))
    			{
    				if(end==-1)
    				{
    					end = i;
    					start = i;
    				}
    				else start = i;
    			}
    		}
    		
    		if(start!=-1 && end!=-1)
    		{
    			if(start!=end)
    			{
    				System.out.print("Please define grades for range: "+ start+"-"+end +" for subject "+name);
    			}
    			else 
    			{
    				System.out.print("Please define grades for mark: "+ start+" for subject "+name); // if a particular mark is missing
    			}
    			System. exit(0);
    		}
    		
    	}
    	
    	
    	static void writeToExcel(ArrayList<StudentDetails> res)throws IOException  //writing to excel files
    	{
    		
    		XSSFWorkbook workbook = new XSSFWorkbook(); 
    		XSSFSheet sheet = workbook.createSheet("Final_Result");
    		
    		int r=0,cn=0; // to change rows and columns
    		
    		Row row=sheet.createRow(r++);
    		
			Cell cell = row.createCell(cn++);
			cell.setCellValue("SL.NO."); 
			cell = row.createCell(cn++);
			cell.setCellValue("NAME");
			cell = row.createCell(cn++);
			cell.setCellValue("SUBJECT");
			cell = row.createCell(cn++);
			cell.setCellValue("MARKS");
			cell = row.createCell(cn++);
			cell.setCellValue("ELIGIBLE FOR GRADE");
			cell = row.createCell(cn++);
			cell.setCellValue("SUBJECT GRADE");
			cell = row.createCell(cn++);
			cell.setCellValue("CREDITS");
			cell = row.createCell(cn++);
			cell.setCellValue("GRADE FINAL POINTS");
			
    		for(StudentDetails s:res)
    		{
    			cn=0;
    		    row=sheet.createRow(r++);
    		    
    		    cell = row.createCell(cn++);
    			cell.setCellValue((int)s.slno);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((String)s.name);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((String)s.sub);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((double)s.marks);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((String)s.eligibility);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((String)s.grade);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((double)s.credit);
    			
    			cell = row.createCell(cn++);
    			cell.setCellValue((double)s.gradeFP);
    		}
    		
    		try
    		  {
    		   
    			for (int i=0; i<8; i++)
    			{
    				sheet.autoSizeColumn(i); // To resize the column
    			}
    			
    		   FileOutputStream out = new FileOutputStream(new File("Final_Result.xlsx"));
    		   
    		   workbook.write(out);
    		   out.close();
    		   
    		   System.out.println("Final_Result.xlsx has been created successfully!");
    		   
    		  } 
    		  catch (Exception e) 
    		  {
    		   e.printStackTrace();
    		  }
    		  finally
    		  {
    		   workbook.close();
    		  }
    		
    	}
    	
    	static Map<String,SubjectDetails> subdetails;// store the details of each subject
    	static ArrayList<StudentDetails> result;    //To store the final result
    	
    	
    	
    	
	    public static void main(String[] args)   
	    {  
	    	
		    // For Marking Scheme
		    try  
		    {  
			    File file = new File("Marking_Scheme.xlsx");   //creating a new file instance  
			    FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file   
			    XSSFWorkbook wb = new XSSFWorkbook(fis);  //creating Workbook instance that refers to .xlsx file   
			    XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			    Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			    
			    subdetails=new HashMap<>();
			    
			    if(itr.hasNext()) // To skip to first row
			    {
			    	itr.next();
			    }
			    
			    while (itr.hasNext())                 
			    {  
				    Row row = itr.next();  
				    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
				    
				    double slno=cellIterator.next().getNumericCellValue();
				    String sub=cellIterator.next().getStringCellValue();
				    String range=cellIterator.next().getStringCellValue();
			    	String grade=cellIterator.next().getStringCellValue();
			    	double credit=cellIterator.next().getNumericCellValue();
			    	double minMarks=cellIterator.next().getNumericCellValue();
			    	double gradePoint=cellIterator.next().getNumericCellValue();
			    	
			    	if(sub.equals("N/A"))
			    	{
			    		System.out.println("Missing Subject Name");
						System. exit(0);	
			    	}
			    	
			    	if(range.equals("N/A"))
			    	{
			    		System.out.println("Missing Range details");
						System. exit(0);	
			    	}
			    	
			    	if(grade.equals("N/A"))
			    	{
			    		System.out.println("Missing Grade details");
						System. exit(0);	
			    	}
			    	
			    	if(credit==0)
			    	{
			    		System.out.println("Missing Credit details");
						System. exit(0);	
			    	}
			    	
			    	if(minMarks==0)
			    	{
			    		System.out.println("Missing Minimum marks");
						System. exit(0);	
			    	}
			    	
			    	if(gradePoint==0)
			    	{
			    		System.out.println("Missing Grade Points");
						System. exit(0);	
			    	}
		
			    	String r[]=range.split("-");
					int start=Integer.parseInt(r[0]);
					int end=Integer.parseInt(r[1]);
					
				    if(subdetails.containsKey(sub)) 
				    {
				    	SubjectDetails sd=subdetails.get(sub);
				    	
		    			for(int i=start;i<=end;i++)
		    			{
		    				sd.grade[i]=grade;
		    				sd.gradePoint[i]=gradePoint;
		    				if(sd.credits!=credit) // To check if the credits entered is same
		    				{
		    					System.out.println("Credits is not matching as per the subject");
		    					System. exit(0);
		    				}
		    				
		    				if(sd.minMarks!=minMarks) //To check if the minimum marks is same
		    				{
		    					System.out.println("Minimum Marks is not matching as per the subject");
		    					System. exit(0);
		    				}
		    			}
		    			subdetails.replace(sub,sd);
				    }
				    else 
				    {
				    	SubjectDetails sd=new SubjectDetails();
				    	for(int i=start;i<=end;i++) 
				    	{
		    				sd.grade[i]=grade;
		    				sd.gradePoint[i]=gradePoint;
		    				sd.credits=credit;
		    				sd.minMarks=minMarks;
		    			}
				    	subdetails.put(sub,sd);
				    }
			    }
			    
			    for(String a:subdetails.keySet()) // checking the range for each subject
			    {
			    	checker(subdetails.get(a).grade,subdetails.get(a).minMarks,a);		
			    }
			    
		 		wb.close();//close the marking scheme file
		 		   
			    //System.out.println("Error free till here!"); 
		    }
		    
		    catch(Exception e)  
		    {  
		    	e.printStackTrace();  
		    }
		    
		    // For Students
		    
		    try  
		    {
		    	
			    File file = new File("Students_Marks_Details.xlsx");   //creating a new file instance  
			    FileInputStream fis = new FileInputStream(file);   //obtaining bytes from the file   
			    XSSFWorkbook wb = new XSSFWorkbook(fis);   //creating Workbook instance
			    XSSFSheet sheet = wb.getSheetAt(0);     //creating a Sheet object to retrieve object  
			    Iterator<Row> itr = sheet.iterator();    //iterating over excel file  
			    
			    result = new ArrayList<>();
			    
			    if(itr.hasNext()) // To skip to first row
			    {
			    	itr.next();
			    }
			    
			    while (itr.hasNext())                 
			    {  
				    Row row = itr.next();  
				    Iterator<Cell> cellIterator = row.cellIterator();   //iterating over each column
				    
				    double Sl=cellIterator.next().getNumericCellValue();
				    String student=cellIterator.next().getStringCellValue();
				    String sub=cellIterator.next().getStringCellValue();
			    	double marks=cellIterator.next().getNumericCellValue();
			    	
			    	if(student.equals("N/A"))
			    	{
			    		System.out.println("Missing Student Name"); // Student Name is N/A
						System. exit(0);	
			    	}
			    	
			    	if(sub.equals("N/A"))
			    	{
			    		System.out.println("Missing Subject Name"); // Subject Name is N/A
						System. exit(0);	
			    	}
			    	
				    if(!subdetails.containsKey(sub)) // To remove the Students without marking Scheme details
				    {
				    	continue;		
				    }
				    
			    	SubjectDetails sd=subdetails.get(sub);
			    	
			    	String grade=sd.grade[(int)Math.ceil(marks)];
			    	String eligibility="NO";
			    	double credit=sd.credits,gradeFP=0;
			    	int slno=result.size()+1;
			    	
			    	if(!grade.equals("F")) 
			    	{
			    		eligibility="YES";
			    		gradeFP=credit*sd.gradePoint[(int)Math.ceil(marks)];
			    	}
			    	
			    	else
			    	{
			    		grade="NO";
			    	}
		    		StudentDetails std=new StudentDetails(slno,marks,credit,gradeFP,student,eligibility,grade,sub);
		    		result.add(std);		 
			    }
			    
			    writeToExcel(result); // write the results to excel file
			    wb.close();//close the marking scheme file
			    
			    
		//	    for(StudentDetails std:result) // Testing
		//	    {
		//	    	System.out.println(std.sl+"\t"+std.name+"\t"+std.sub+"\t"+std.marks+"\t"+std.eligibility+"\t"+std.grade+"\t"+std.credit+"\t"+std.gradeFP);
		//	    }
			   
		    }
		    
		    catch(Exception e)  
		    {  
		    e.printStackTrace();  
		    } 
		    
		    
	    }  
    }  