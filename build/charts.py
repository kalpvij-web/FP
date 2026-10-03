import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
plt.rcParams["font.family"]="serif"; plt.rcParams["font.serif"]=["Liberation Serif","Times New Roman","DejaVu Serif"]
def bar(labels, vals, title, xlabel, fn, horiz=False):
    fig, ax = plt.subplots(figsize=(6.4,3.4), dpi=200)
    if horiz:
        y=range(len(labels)); ax.barh(list(y), vals, color="#3b5b8c"); ax.set_yticks(list(y)); ax.set_yticklabels(labels); ax.invert_yaxis(); ax.set_xlabel(xlabel)
        for i,v in enumerate(vals): ax.text(v+0.05,i,str(v),va="center",fontsize=10)
        ax.set_xlim(0,max(vals)+1)
    else:
        ax.bar(labels, vals, color="#3b5b8c", width=0.55); ax.set_ylabel(xlabel)
        for i,v in enumerate(vals): ax.text(i,v+0.15,str(v),ha="center",fontsize=10)
        ax.set_ylim(0,max(vals)+1.5)
    for s in ["top","right"]: ax.spines[s].set_visible(False)
    ax.set_title(title, fontsize=11)
    fig.tight_layout(); fig.savefig(fn); plt.close(fig)
bar(["Pinterest","Instagram"],[5,2],"Posters in the corpus, by platform","Number of posters","g1.png")
bar(["Intertextuality","Irony","Parody","Wordplay / pun","Comparative satire","Slang reclamation","Dark humor"],[3,3,3,2,1,1,1],"Rhetorical devices in the 7 analysed items","Number of items using the device","g2.png",horiz=True)
bar(["Absence of\naccountability","Systemic\ncollapse","Delay and\nsilence","Exam security\n(paper leak)","Stereotype of\napolitical youth"],[3,1,1,1,1],"What the 7 analysed items were targeting","Number of items","g3.png")
