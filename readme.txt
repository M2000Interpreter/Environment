M2000 Interpreter and Environment
Version 13 revision 45 active-X

1. Optimizations

2. We can put groups in arrays of objects of type RefArray. We can call methods using dot or => (always we put a pointer)


Example:
I Use for Tree an array of objects (nodes) (the eertree example in info use inventory, a map list). We can do that because all indexes is 0,1,2... like array indexes. 

eertree= lambda (s as string)
	->{
		Class Node {
			inventory myedges
			long length, suffix=0
			Function edges(s as string) {
				=-1 : if exist(.myedges, s) then =eval(.myedges)
			}
			Module edges_append (a as string, where as long) {
				Append .myedges, a:=where
			}
			Class:
			Module Node(.length) {
				Read ? .suffix, .myedges
			}
		}	
		object Tree[100]
		' we can give 1, Tree[1] and this type add items as we set index above upper limit
		Tree[0]=Node(0,1)
		Tree[1]=Node(-1,1)
		k=0
		suffix=0
		for i=0 to len(s)-1
			d=mid$(s,i+1,1)
			n=suffix
			Do
				k=Tree[n].length
				b=i-k-1
				if b>=0 then if mid$(s,b+1,1)=d Then exit
				n =Tree[n].suffix  
			Always
			e=Tree[n].edges(d)
			if e>=0 then suffix=e :continue
			suffix=len(Tree)
			Tree[len(Tree)]=Node(k+2)
			Tree[n].edges_append d, suffix
			If tree[suffix].length=1 then tree[suffix].suffix=0: continue
			Do
				n=Tree[n].suffix
				b=i-Tree[n].length-1
				if b>0 Then If  mid$(s, b+1,1)=d then exit
			Always
			e=Tree[n].edges(d)
			if e>=0 then tree[suffix].suffix=e
		next
		=tree
	}
children=lambda (s as array, Tree,  n, root as string="")
	-> {
	recur(n, root)
	=s
	sub recur(n, root)
		local Long L=Len(Tree[n].myEdges), i, c, nxt
		local	String d, p	
		if L=0 then exit sub
		do	c=Tree[n].myEdges
			d=Eval$(c, i)  ' read keys at position i
			nxt=c(i!)   '  read value using position 
			p = if(n=1 -> d, d+root+d)
			append s, (p,)
			recur(nxt, p)
			i++
		when i<L
	end sub
	}
Palindromes=lambda children (Tree as *Object[])->"("+quote$(children(children((,), Tree, 0), Tree, 1)#str$({", "}))+")"
Print Palindromes(eertree("987654321eertree12345678954321eertree12345eertree"))
Print Palindromes(eertree("banana"))


George Karras, Kallithea Attikis, Greece.
fotodigitallab@gmail.com

The first time Windows make some work behind the scene so the M200 console slow down. So type END and open it again.

To get the INFO file, from M2000 console do these:
dir appdir$
load info
then press F1 to save info.gsb to M2000 user directory

You can also execute statement SETTINGS to change font/language/colors and size of console letters.

Read wiki at Github for compiling M2000 from source.

From version 9.0 revision 50:
there is a new ca.crt - install ca.crt as root certificate (optional)
install ca.crt as root certificate (optional).


http://georgekarras.blogspot.gr/

https://rosettacode.org/wiki/Category:M2000_Interpreter (544 tasks)

Code/Exe files can be found here: 

https://github.com/M2000Interpreter                 