# Import
import math #math
import pandas as pd #to manipulate data
import os #to access system directories and files
import time #to measure run time
import itertools #for combination and permutation

# Constants
XLSXNAME = "spcol-batcher-input.xlsx"
XLSXWALLS = "wall_info"
XLSXETABS = "ETABS"
OUTPUT_FOLDER = "spcol-batcher-output"
REBAR = {
    3:[0.375, 0.11],
    4:[0.500, 0.20],
    5:[0.625, 0.31],
    6:[0.750, 0.44],
    7:[0.875, 0.60],
    8:[1.000, 0.79],
    9:[1.128, 1.00],
    10:[1.270, 1.27],
    11:[1.410, 1.56],
    14:[1.693, 2.25],
    18:[2.257, 4.00]}




"""
############################################
HELPER FUNCTIONS
############################################
"""

def trim_dataframe1(df_main):
    """
    initial trim of imported data
    """
    generalinfo = []
    generalinfo.append(df_main.iloc[14,1]) #0 = project name
    generalinfo.append(df_main.iloc[15,1]) #1 = engineer
    generalinfo.append(df_main.iloc[16,1]) #2 = f'c
    generalinfo.append(df_main.iloc[17,1]) #3 = ec
    generalinfo.append(df_main.iloc[18,1]) #4 = 0.85f'c
    generalinfo.append(df_main.iloc[19,1]) #5 = beta
    generalinfo.append(df_main.iloc[20,1]) #6 = ecu
    generalinfo.append(df_main.iloc[21,1]) #7 = fy
    generalinfo.append(df_main.iloc[22,1]) #8 = Es
    generalinfo.append(df_main.iloc[23,1]) #9 = ey
    generalinfo.append(df_main.iloc[24,1]) #10 = cb
    
    # load combination info
    LC = list(df_main.iloc[28:51,1].dropna())
    
    # column info
    header = df_main.iloc[54]
    N_rows = df_main.iloc[56:,:].count()[0]
    df_wall = pd.DataFrame(df_main.iloc[56:(56+N_rows),:])
    df_wall.columns = header
    df_wall = df_wall.reset_index(drop=True)
    return generalinfo, LC, df_wall


def create_output_folder(output_folder):
    """
    make directory for output files
    """
    parent_dir = os.getcwd()
    if os.path.isdir(output_folder):
        folderexists = output_folder
        j=1
        while os.path.isdir(output_folder):
            output_folder = folderexists + str(j)
            j=j+1
        output_dir = os.path.join(parent_dir,output_folder)
        os.makedirs(output_dir)
        print("{} folder already exists, generating files in {}...".format(folderexists,output_folder))
    else:
        output_dir = os.path.join(parent_dir,output_folder)
        os.makedirs(output_dir)
        print("generating files in {}...".format(output_folder))   
    
    os.makedirs(os.path.join(output_dir,"importfiles"))
    os.makedirs(os.path.join(output_dir,"csv"))
    return output_dir


def generate_pmv(LC,df_wall):
    """
    for every section, generate all its associated P,M,V,ID values
    ID is used to identify wall and mix of load combos
    """
    PMV_dict={}
    for i in range(len(df_wall)):
        pier = df_wall.iloc[i,0]
        PMV_dict[pier]=[]
        
        for loadcombo in LC:
            pmv=list(df_ETABS.query('Pier == "{}" and Output_Case == "{}" and Location=="Bottom"'.format(
            pier,loadcombo))[["P","V2","M3"]].iloc[0])
            pmv[0]=-pmv[0]/1000
            pmv[1]=pmv[1]/1000
            pmv[2]=pmv[2]/12000
            ID = pier+"_"+loadcombo
            pmv.append(ID)
            PMV_dict[pier].append(pmv)
                  
    return PMV_dict


def draw_bars(x,y,b,h,cb,nx,ny,A,D,layout):
    """
    This function generates the x,y coordinate of boundary element steel based on
    the following input:
    x : dx from origin to lower left of boundary element
    y : dy from origin to lower left of boundary element
    b : dimension of BE along x
    h : dimension of BE along y
    cb : clear cover to face of rebar
    nx : number of columns along x
    ny : number of columns along y
    A : area of rebar
    D : diameter of rebar
    layout : layout. "full" or "perimeter"
    """
    # first determine spacing and coordinate with local origin at lower left
    cx = b - 2*cb - D
    cy = h - 2*cb - D
    sx = cx / (nx-1)
    sy = cy / (ny-1)
    xcoord=[]
    ycoord=[]
    xcoord.append(cb+D/2)
    ycoord.append(cb+D/2)
    for i in range(nx-1):
        xcoord.append(xcoord[-1]+sx)
    for i in range(ny-1):
        ycoord.append(ycoord[-1]+sy)
    reinfarray_local = list(itertools.product(xcoord,ycoord))
    
    if layout == "perimeter":
        x0 = cb+D/2
        xn = xcoord[-1]
        y0 = cb+D/2
        yn = ycoord[-1]
        reinfarray_local = [e for e in reinfarray_local if e[0]==x0 or e[0]==xn or e[1]==y0 or e[1]==yn]
            
    # translate local origin to global
    reinfarray_global=[]
    for i in range(len(reinfarray_local)):
        x_global = reinfarray_local[i][0] + x 
        y_global = reinfarray_local[i][1] + y 
        reinfarray_global.append([A,x_global,y_global])
        
    return reinfarray_global


def draw_web_bars(x,y,b,h,cb,s_target,A,D,ncurtain,direction):
    """
    This function generates the x,y coordinate of boundary element steel based on
    the following input:
    x : dx from origin to lower left corner of web
    y : dy from origin to lower left corner of web
    b : dimension along x
    h : dimension along y
    cb : clear cover to face of rebar
    s_target : target spacing of bars
    A : area of rebar
    D : diameter of rebar
    ncurtain: number of curtains
    direction: vertical or horizontal
    """
    xcoord=[]
    ycoord=[]
    reinfarray_local=[]
    if direction == "v":
        s_t = (b-cb-cb-D)/(ncurtain-1)
        N_L = math.floor((h)/s_target)
        s_L = (h)/(N_L+1)
        xcoord.append(cb+D/2)
        ycoord.append(s_L)
        for i in range(ncurtain-1):
            xcoord.append(xcoord[-1]+s_t)
        for i in range(N_L-1):
            ycoord.append(ycoord[-1]+s_L)
        reinfarray_local = list(itertools.product(xcoord,ycoord))
    elif direction == "h":
        s_t = (h-cb-cb-D)/(ncurtain-1)
        N_L = math.floor((b)/s_target)
        s_L = (b) / (N_L+1)
        ycoord.append(cb+D/2)
        xcoord.append(s_L)
        for i in range(ncurtain-1):
            ycoord.append(ycoord[-1]+s_t)
        for i in range(N_L-1):
            xcoord.append(xcoord[-1]+s_L)
        reinfarray_local = list(itertools.product(xcoord,ycoord))

    # translate local origin to global
    reinfarray_global=[]
    for i in range(len(reinfarray_local)):
        x_global = reinfarray_local[i][0] + x 
        y_global = reinfarray_local[i][1] + y 
        reinfarray_global.append([A,x_global,y_global])
    return reinfarray_global


def get_pts_column(rowdata):
    """
    function to generate external points if type is a column
    """
    b=rowdata[2]
    h=rowdata[3]
    N=5
    ptarray=[
        [-b/2,-h/2],
        [b/2,-h/2],
        [b/2,h/2],
        [-b/2,h/2],
        [-b/2,-h/2]]
    return ptarray, N-1


def get_reinf_column(rowdata,cb):
    """
    function to generate external points if type is a column
    """
    b=rowdata[2]
    h=rowdata[3]
    nx = rowdata[19]
    ny = rowdata[20]
    barsize = rowdata[21]
    layout = rowdata[23]
    A = REBAR[barsize][1]
    D = REBAR[barsize][0]
    reinfarray = draw_bars(-b/2,-h/2,b,h,cb,nx,ny,A,D,layout)
    if layout == "full":
        N = nx * ny
    elif layout == "perimeter":
        N = nx*2 + (ny-2)*2
    else:
        raise RuntimeError('boundary element layout type not recognized')
    return reinfarray, N


def get_pts_wall(rowdata):
    """
    function to generate external points if type is a wall
    """
    bw=rowdata[2]
    Lw=rowdata[3]
    tfb=rowdata[4]
    tft=rowdata[5]
    bb1=rowdata[6]
    bb2=rowdata[7]
    bt1=rowdata[8]
    bt2=rowdata[9]
    ptarray_copy = [
        [-bw/2, -Lw/2],
        [-bw/2-bb1, -Lw/2],
        [-bw/2-bb1, -Lw/2-tfb],
        [bw/2+bb2, -Lw/2-tfb],
        [bw/2+bb2, -Lw/2],
        [bw/2, -Lw/2],
        [bw/2, Lw/2],
        [bw/2+bt2, Lw/2],
        [bw/2+bt2, Lw/2+tft],
        [-bw/2-bt1, Lw/2+tft],
        [-bw/2-bt1, Lw/2],
        [-bw/2, Lw/2],
        [-bw/2, -Lw/2]]
    # need to remove duplicate from list. set() does not preserve order
    ptarray = []
    for i in ptarray_copy:
        if i not in ptarray:
            ptarray.append(i)
    ptarray.append(ptarray_copy[0])
    return ptarray, len(ptarray)-1


def get_reinf_wall(rowdata,cb):
    """
    function to generate reinforcement coordinates if type is a wall
    """
    bw=rowdata[2]
    Lw=rowdata[3]
    tfb=rowdata[4]
    tft=rowdata[5]
    bb1=rowdata[6]
    bb2=rowdata[7]
    bt1=rowdata[8]
    bt2=rowdata[9]
    nx=[rowdata[19],rowdata[24],rowdata[29],rowdata[34],rowdata[39],rowdata[44]]
    ny=[rowdata[20],rowdata[25],rowdata[30],rowdata[35],rowdata[40],rowdata[45]]
    barsize=[rowdata[21],rowdata[26],rowdata[31],rowdata[36],rowdata[41],rowdata[46]]
    Lbe=[rowdata[22],rowdata[27],rowdata[32],rowdata[37],rowdata[42],rowdata[47]]
    layout=[rowdata[23],rowdata[28],rowdata[33],rowdata[38],rowdata[43],rowdata[48]]
    
    # determine which BE needs to be drawn
    BE_flag = [1,1,1,1,1,1]
    for i in range(6):
        if math.isnan(nx[i]) or math.isnan(ny[i]):
            BE_flag[i]=0
    if tfb == 0:
        BE_flag[2]=0
        BE_flag[3]=0
    if tft == 0:
        BE_flag[4]=0
        BE_flag[5]=0
        
    # draw BE steel
    BE_reinf=[]
    if BE_flag[0]==1:
        dx = -bw/2
        dy = -Lw/2-tfb
        b = bw
        h = Lbe[0]
        A = REBAR[barsize[0]][1]
        D = REBAR[barsize[0]][0]
        BE_reinf.append(draw_bars(dx,dy,b,h,cb,nx[0],ny[0],A,D,layout[0]))
    if BE_flag[1]==1:
        dx = -bw/2
        dy = Lw/2 + tft - Lbe[1]
        b = bw
        h = Lbe[1]
        A = REBAR[barsize[1]][1]
        D = REBAR[barsize[1]][0]
        BE_reinf.append(draw_bars(dx,dy,b,h,cb,nx[1],ny[1],A,D,layout[1]))
    if BE_flag[2]==1:
        dx = -bw/2 - bb1
        dy = -Lw/2 - tfb
        b = Lbe[2]
        h = tfb
        A = REBAR[barsize[2]][1]
        D = REBAR[barsize[2]][0]
        BE_reinf.append(draw_bars(dx,dy,b,h,cb,nx[2],ny[2],A,D,layout[2]))
    if BE_flag[3]==1:
        dx = bw/2 + bb2 - Lbe[3]
        dy = -Lw/2 - tfb
        b = Lbe[3]
        h = tfb
        A = REBAR[barsize[3]][1]
        D = REBAR[barsize[3]][0]
        BE_reinf.append(draw_bars(dx,dy,b,h,cb,nx[3],ny[3],A,D,layout[3]))
    if BE_flag[4]==1:
        dx = -bw/2 - bt1
        dy = Lw/2
        b = Lbe[4]
        h = tft
        A = REBAR[barsize[4]][1]
        D = REBAR[barsize[4]][0]
        BE_reinf.append(draw_bars(dx,dy,b,h,cb,nx[4],ny[4],A,D,layout[4]))
    if BE_flag[5]==1:
        dx = bw/2 + bt2 - Lbe[5]
        dy = Lw/2
        b = Lbe[5]
        h = tft
        A = REBAR[barsize[5]][1]
        D = REBAR[barsize[5]][0]
        BE_reinf.append(draw_bars(dx,dy,b,h,cb,nx[5],ny[5],A,D,layout[5]))
    
    # Web steel(0 = bottom left, 1 = bottom right, 2 = web, 3 = top left, 4 = top right)
    ncurtain=[rowdata[13],rowdata[13],rowdata[10],rowdata[16],rowdata[16]]
    s_target=[rowdata[15],rowdata[15],rowdata[12],rowdata[18],rowdata[18]]
    web_flag = [1,1,1,1,1] 
    if i in range(5):
        if math.isnan(ncurtain[i]) or math.isnan(s_target[i]):
            web_flag[i]=0
    if tfb==0 or bb1==0:
        web_flag[0]=0
    if tfb==0 or bb2==0:
        web_flag[1]=0
    if tft==0 or bt1==0:
        web_flag[3]=0
    if tft==0 or bt2==0:
        web_flag[4]=0
    
    # if LBE is left blank as Nan, encounters error. Change them to 0
    for i in range(6):
        index = 22+5*i
        if math.isnan(rowdata[index]):
            rowdata[index]=0
    
    # Draw web steel
    web_reinf=[]
    if web_flag[0]==1:
        dx = -bw/2 - bb1 + rowdata[32]
        dy = -Lw/2 - tfb
        b =  bb1 - rowdata[32]
        h = tfb
        s = s_target[0]
        A = REBAR[rowdata[14]][1]
        D = REBAR[rowdata[14]][0]
        web_reinf.append(draw_web_bars(dx,dy,b,h,cb,s,A,D,ncurtain[0],'h'))
    if web_flag[1]==1:
        dx = bw/2
        dy = -Lw/2 - tfb
        b = bb2 - rowdata[37]
        h = tfb
        s = s_target[1]
        A = REBAR[rowdata[14]][1]
        D = REBAR[rowdata[14]][0]
        web_reinf.append(draw_web_bars(dx,dy,b,h,cb,s,A,D,ncurtain[1],'h'))
    if web_flag[2]==1:
        dx = -bw/2
        dy = -Lw/2 - tfb + rowdata[22]
        b = bw
        h = Lw + tfb + tft - rowdata[22] - rowdata[27]
        s = s_target[2]
        A = REBAR[rowdata[11]][1]
        D = REBAR[rowdata[11]][0]
        web_reinf.append(draw_web_bars(dx,dy,b,h,cb,s,A,D,ncurtain[2],'v'))
    if web_flag[3]==1:
        dx = -bw/2 - bt1 + rowdata[42]
        dy = Lw/2
        b = bt1 - rowdata[42]
        h = tft
        s = s_target[3]
        A = REBAR[rowdata[17]][1]
        D = REBAR[rowdata[17]][0]
        web_reinf.append(draw_web_bars(dx,dy,b,h,cb,s,A,D,ncurtain[3],'h'))
    if web_flag[4]==1:
        dx = bw/2
        dy = Lw/2
        b = bt2 - rowdata[47]
        h = tft
        s = s_target[4]
        A = REBAR[rowdata[17]][1]
        D = REBAR[rowdata[17]][0]
        web_reinf.append(draw_web_bars(dx,dy,b,h,cb,s,A,D,ncurtain[4],'h'))
    
    reinfarray_copy = BE_reinf + web_reinf
    reinfarray=[]
    for i in reinfarray_copy:
        for j in i:
            reinfarray.append(j)
            
    return reinfarray, len(reinfarray)


def create_import_files(output_dir,rowdata,externalpts,N_pts,reinf,N_reinf,PMV):
    """
    Generate importable .txt files within a "importfiles" subdirectory. For 
    reference purposes if using import feature is more desireable
    """
    filename = os.path.join(output_dir, "importfiles","GEOMETRY_"+rowdata[0]+".txt")
    with open(filename,'w') as f:
        f.write("{}\t\n".format(N_pts))
        for j in range(len(externalpts)-1):
            items = externalpts[j]
            f.write("{}\t{}\n".format(items[0],items[1]))
        f.write("0")
    filename2 = os.path.join(output_dir, "importfiles","REBAR_"+rowdata[0]+".txt")
    with open(filename2,'w') as f2:
        f2.write("{}\t\t\n".format(N_reinf))
        for items in reinf:
            f2.write("{}\t{}\t{}\n".format(items[0],items[1],items[2]))
    filename3 = os.path.join(output_dir, "importfiles","LOAD_"+rowdata[0]+".txt")
    with open(filename3,'w') as f3:
        f3.write("{}\t\t\n".format(len(PMV)))
        for items in PMV:
            f3.write("{}\t{}\t{}\n".format(items[0],items[2],0))
            
            
def create_cti_files(output_dir,rowdata,externalpts,N_pts,reinf,N_reinf,PMV,generalinfo):
    """
    Generate the spColumn .cti files
    """
    filename = os.path.join(output_dir,rowdata[0]+".cti")
    with open(filename,'w') as f:
        f.write("#spColumn Text Input (CTI) File\n")
        f.write("[spColumn Version]\n")
        f.write("6.000\n")
        f.write("[Project]\n")
        f.write(generalinfo[0]+"\n")
        f.write("[Column ID]\n")
        f.write(rowdata[0]+"\n")
        f.write("[Engineer]\n")
        f.write(generalinfo[1]+"\n")
        f.write("[Investigation Run Flag]\n")
        f.write("15\n")
        f.write("[Design Run Flag]\n")
        f.write("9\n")
        f.write("[Slenderness Flag]\n")
        f.write("0\n")
        f.write("[User Options]\n")
        f.write("0,0,6,0,0,0,0,0,2,0,0,0,0,-1,3,-1,{},{},0,{},0,0,0,0,0,13\n".format(N_reinf,len(PMV),N_pts+1))
        f.write("[Irregular Options]\n")
        f.write("-2,0,0,1,0.6,50,50,-50,-50,5,5,2.5,2.5\n")
        f.write("[Ties]\n")
        f.write("1,1,7\n")
        f.write("[Investigation Reinforcement]\n")
        f.write("0,{},0,0,0,0,0,0,0,0,0,0\n".format(str(math.floor(N_reinf/2))))
        f.write("[Design Reinforcement]\n")
        f.write("0,0,0,0,0,0,0,0,0,0,0,0\n")
        f.write("[Investigation Section Dimensions]\n")
        f.write("0,0\n")
        f.write("[Design Section Dimensions]\n")
        f.write("0,0,0,0,0,0\n")
        f.write("[Material Properties]\n")
        f.write("{:.6f},{:.6f},{:.6f},{:.6f},{:.6f},{:.6f},{:.6f},0,1,1,{:.6f}\n".format(
            generalinfo[2],generalinfo[3],generalinfo[4],generalinfo[5],generalinfo[6],
            generalinfo[7],generalinfo[8],generalinfo[9]))
        f.write("[Reduction Factors]\n")
        f.write("0.8,0.9,0.65,0.1,0\n")
        f.write("[Design Criteria]\n")
        f.write("0.01,0.08,1.5,1\n")
        f.write("[External Points]\n")
        f.write(str(N_pts+1)+"\n")
        for pt in externalpts:
            f.write("{:.6f},{:.6f}\n".format(pt[0],pt[1]))
        f.write("[Internal Points]\n")
        f.write("0\n")
        f.write("[Reinforcement Bars]\n")
        f.write(str(N_reinf)+"\n")
        for re in reinf:
            f.write("{:.6f},{:.6f},{:.6f}\n".format(re[0],re[1],re[2]))
        f.write("[Factored Loads]\n")
        f.write(str(len(PMV))+"\n")
        for load in PMV:
            f.write("{:.6f},{:.6f},{}\n".format(load[0],load[2],"0"))
        f.write("[Slenderness: Column]\n")
        f.write("0,0,0,1,0,1,1,0\n")
        f.write("0,0,0,1,0,1,1,0\n")
        f.write("[Slenderness: Column Above And Below]\n")
        f.write("1,0,0,0,{:.6f},{:.6f}\n".format(generalinfo[2],generalinfo[3]))
        f.write("1,0,0,0,{:.6f},{:.6f}\n".format(generalinfo[2],generalinfo[3]))
        f.write("[Slenderness: Beams]\n")
        for i in range(8):
            f.write("1,0,0,0,0,{:.6f},{:.6f}\n".format(generalinfo[2],generalinfo[3]))
        f.write("[EI]\n")
        f.write("0\n")
        f.write("[SldOptFact]\n")
        f.write("0\n")
        f.write("[Phi_Delta]\n")
        f.write("0.75\n")
        f.write("[Cracked I]\n")
        f.write("0.35,0.75\n")
        f.write("[Service Loads]\n")
        f.write("0\n")
        f.write("[Load Combinations]\n")
        f.write("13\n")
        f.write("1.4,0,0,0,0\n")
        f.write("1.2,1.6,0,0,0.5\n")
        f.write("1.2,1,0,0,1.6\n")
        f.write("1.2,0,0.8,0,1.6\n")
        f.write("1.2,1,1.6,0,0.5\n")
        f.write("0.9,0,1.6,0,0\n")
        f.write("1.2,0,-0.8,0,1.6\n")
        f.write("1.2,1,-1.6,0,0.5\n")
        f.write("0.9,0,-1.6,0,0\n")
        f.write("1.2,1,0,1,0.2\n")
        f.write("0.9,0,0,1,0\n")
        f.write("1.2,1,0,-1,0.2\n")
        f.write("0.9,0,0,-1,0\n")
        f.write("[BarGroupType]\n")
        f.write("1\n")
        f.write("[User Defined Bars]\n")
        f.write("[Sustained Load Factors]\n")
        f.write("100,0,0,0,0\n")
        
        
def create_batch_file(output_dir):
    """
    Generate .bat file to batch analyze .cti files
    """
    filename = os.path.join(output_dir,"batch_run_analysis.bat")
    with open(filename,'w') as f:
        for i in range(len(df_wall)):
            f.write("spColumn /i:{}.cti /o:{}.out /emf\n".format(df_wall.iloc[i,0], str(df_wall.iloc[i,0])+"_result"))


def pickle_data(output_dir,df_wall,df_ETABS,PMV_dict):
    """
    Save some data to csv file in case needed when post-processing results
    """
    filename1 = os.path.join(output_dir,"csv","walldata.csv")
    filename2 = os.path.join(output_dir,"csv","ETABS.csv")
    filename3 = os.path.join(output_dir,"csv","PMV.csv")
    df_wall.to_csv(filename1, index=False)
    df_ETABS.to_csv(filename2, index=False)
    
    with open(filename3,'w') as f:
        f.write("ID,P(kips),V2(kips),M3(kip.ft)\n")
        for key in PMV_dict:
            value = PMV_dict[key]
            for items in value:
                f.write("{},{},{},{}\n".format(items[3],items[0],items[1],items[2]))
    
    

"""
############################################
MAIN
############################################
"""
time_start = time.time()
# 1.) Extract dataframe from Excel
df_main = pd.read_excel(XLSXNAME, sheet_name=XLSXWALLS, usecols='A:AW')
df_ETABS = pd.read_excel(XLSXNAME, sheet_name=XLSXETABS, usecols='A:K')
df_ETABS.columns=[column.replace(" ", "_") for column in df_ETABS.columns] 


# 2.) Trim data
generalinfo, LC, df_wall = trim_dataframe1(df_main)
print("Input file read successfully with {} rows...".format(len(df_wall)))
print("Generating .cti files. Expected wait time about 10s to 60s...")

# 3.) Create output folder
output_dir = create_output_folder(OUTPUT_FOLDER)


# 4.) Organize load combinations
PMV_dict = generate_pmv(LC,df_wall)


# 5.) Loop through every section
for i in range(len(df_wall)):
    rowdata = df_wall.iloc[i]
    # 5a.) Create array for section external points and reinforcement
    if rowdata[1]=="Column":
        externalpts, N_pts = get_pts_column(rowdata)
        reinf, N_reinf = get_reinf_column(rowdata,generalinfo[10])
    elif rowdata[1]=="Wall":
        externalpts, N_pts = get_pts_wall(rowdata)
        reinf, N_reinf = get_reinf_wall(rowdata,generalinfo[10])
    else:
        raise RuntimeError('Type not recognized')


    # 5b.) Create import .txt files and spColumn .cti files
    PMV = PMV_dict[rowdata[0]]
    create_import_files(output_dir,rowdata,externalpts,N_pts,reinf,N_reinf,PMV)
    create_cti_files(output_dir,rowdata,externalpts,N_pts,reinf,N_reinf,PMV,generalinfo)
    
    

# 6.) pickle df_wall, df_ETABS, and PMV points for reference
pickle_data(output_dir,df_wall,df_ETABS,PMV_dict)
    
    
# 7.) Create batch file to run analysis
create_batch_file(output_dir)
print("all spColumn .cti files have been generated. Open .bat file in output folder to batch run analysis")
time_end = time.time()
print("Script finished running. Total elasped time: {:.2f} seconds".format(time_end - time_start))


#if __name__ == "__main__":
#    main()





