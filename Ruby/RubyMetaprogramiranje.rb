require "google_drive"

#Nije radilo drugacije promeni absolutnu u relativnu putanju
session = GoogleDrive::Session.from_config("config.json")

ws = session.spreadsheet_by_key("1py-0UkzfXdYlBNbe73isb2YFnxosTXeSdg4ahJhqDA4").worksheets[0]
ws2 = session.spreadsheet_by_key("1py-0UkzfXdYlBNbe73isb2YFnxosTXeSdg4ahJhqDA4").worksheets[1] #isti dokument, drugi ws
module Oh
    @ws

    def []=(index,arg)
       
        @@ws[index+@@ws.row0,@@ws.headers.index(self[0])+@@ws.col0]=arg if (not @@ws.headers.index(self[0]).nil? and not self[index].nil?)
       original_arr(index,arg)
        
    end
end

class GoogleDrive::Worksheet 
  include Enumerable
  attr_accessor :col0,:row0
  attr_accessor :headers
  attr_accessor :table
  alias :original_arr :[]

  def method_missing(name, *args)
    x=name.to_s
    x=x.chop if x =~ (/=$/)
    
    super unless headers.include?x
    ind=headers.index(name)
    name = name.to_s
    if name =~ (/=$/)
      #TODO 
      instance_variable_set("@#{name.chop}", args.first)
    else
      instance_variable_set("@#{name}", self[name])
      instance_variable_get("@#{name}")
    end
  end




  def[](*args)
  
   
    case args.size
    when 1
      a=Array.new()
      a.extend(Oh)
      Oh.class_variable_set(:@@ws,self)
      if headers.include?args[0]
        (row0..num_rows).each do |row| 
        a[row-row0]=self.original_arr(row,col0+headers.index(args[0]))
       # p self.original_arr(row,col0+headers.index(args[0]))
        end
        return a
      end
    when 2
    return self.original_arr(args[0],args[1])
    end
    
  end

  def row(obj)
    a=[]
   (col0..num_cols).each do |col|
      a[col-col0]=self[obj+row0,col]
    end
   return a 
  end
  def each()
    (row0..num_rows).each do |row|
      (col0..num_cols).each do |col|
        yield self[row,col]
      end
    end
  end
  def find_table()  #trazi pocetak table u ws
    (1..num_rows).each do |row|
      (1..num_cols).each do |col|
        if self[row,col]!=""
          @row0=row
          @col0=col
          return
        end
      end
    end
  end
  def build_table #u principu ne previse korisno ali moze biti
    @table=Array.new(num_rows-row0+1){Array.new(num_cols-col0+1)}
    (row0..num_rows).each do |row|
      (col0..num_cols).each do |col|
        @table[row-row0][col-col0]=self[row,col]
      end
    end
  end
  def print_table()
    (0..table.length-1).each do |row|
     p table[row]
     end
    end
  def find_header()  #Nalazi headere radi lakseg snalazenja (pocetni red)
    a=Array.new
    (col0..num_cols).each do |col|
     a[col-col0]=self[row0, col] unless self[row0,col].nil?
      end
   @headers=a
  end

  def find_total() #provera ukoliko postoji poslednji red kao total
    if self[num_rows,col0]=="total" or self[num_rows,col0]=="subtotal"
        @num_cols=@num_cols-1
        @num_rows=@num_rows-1
        end
    end

    def +(arg) #dodaje redove u ws1 ukoliko imaju iste headere
        x=num_rows
        return 0 unless arg.headers==self.headers
        (x+1..x+arg.num_rows-arg.row0-1).each do |row|
            (col0..num_cols).each  do |col|
                self[row,col]=arg[row-x+arg.row0,arg.col0+col-col0]
            end
        end
        self.save
    end

    def -(arg)#brise redove ukoliko se oba reda nalaze u obe tabele brise iz ws1
      x=num_rows
      return 0 unless arg.headers==self.headers
      (arg.row0+1..arg.num_rows).each do |row|
        (row0..num_rows).each do |nrow|
          self.delete_row(table.index(arg.row(row-arg.row0)))  if self.table.include?(arg.row(row-arg.row0))
        end 
      end
    self.save
  end

  def delete_row(arg) #Pomocna funkcija za -(arg)
    (col0..num_cols).each do |col|
      self[arg+row0,col]=""
    end
    build_table #very bad ispravi nekad
  end



end

#metode koje treba da se zovu kako bi dobili neke potrbene podatke
#BITNO BITNO BITNO BITNO BITNO BITNO BITNO
ws.find_table
ws.find_header
ws.build_table
ws.find_total
ws2.find_table
ws2.find_header
#BITNO BITNO BITNO BITNO BITNO BITNO BITNO


class Array
  attr_accessor :ws
  attr_accessor :headers
  alias :original_sum :sum
  alias :original_plus :+
  alias :original_map :map
  alias :original_reduce :reduce
  alias :origianl_select :select
  alias :original_arr :[]=
  #tehnicki je ovde suma,avg  i to 
  def sum() 
    a=Array.new
  
    x=0
   self.each do |i|
    a[x]=i.to_i
    x=x+1
   end
  
   return a.original_sum

  end

  def avg()
 a=Array.new
 x=0
 self.each do |i|
  a[x]=i.to_i
  x=x+1
 end
   x= a.original_sum
   return x/(length-1)
  end


  def+(arg)
 
   return original_plus(arg) if arg.instance_of?Numeric
    return original_plus(arg.to_i)
 
  end

  def method_missing(name, *args)
    x=name.to_s
    x=x.chop if x =~ (/=$/)
    x=x[1..x.length] if (x=~/[0-9]$/ and x[0]=='n') #specijalno za brojeve
    super unless self.include?x
    name = name.to_s
    ind=self.index(x)

    if name =~ (/=$/)
      #TODO 
      instance_variable_set("@#{name.chop}", args.first)
    else
      instance_variable_set("@#{name}", @@ws.row(ind))
      instance_variable_get("@#{name}")
    end
  end



end
Array.class_variable_set(:@@ws, ws)
#p ws.headers
#p ws.table
#s.print_table
#p ws["StudentID"] 

#p ws["StudentID"][1]
#ws["StudentID"][1]="2"
#ws["StudentID"][1]="1" 

#p ws.row(1)


#p ws.StudentID

#p ws.StudentID.sum

#p ws.StudentID.avg

#p ws.StudentID.n10 #za brojeve

#p ws.FirstName.Jessica

#p ws.StudentID.select{|num| p num if num=="5"}


#p ws.StudentID.map { |cell|  cell.to_i}


 #p ws.StudentID.reduce{ |sum, n|sum.to_i + n.to_i } 

#p ws.row(4)[5] 

#ws.each do |i| 
#  p i

#end
#ws+ws2
#ws-ws2
ws.save
ws.reload
# Dumps all cells.

#p ws-ws2






#end