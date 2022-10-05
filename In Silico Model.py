from brian2 import *
import numpy as np
import matplotlib as mpl
import matplotlib.pyplot as plt
from matplotlib import colors

V_th = 30 * mV  # 1 reset voltage
V_r = -60 * mV  # Voltage to reset to

E_l = -60 * mV  # leak, excitatory, and inhibitory potentials
E_e = 140 * mV
E_i = -80 * mV
w_e = 77 * nsiemens  # excitatory and inhibitory synapse weights
w_i = 20 * psiemens
tau_i = 5 * ms  # time constant
tau_e = 20 * ms
tau_r = 2 * ms  # refractory time
tau_inj = 5* ms  # current injection decay constant
C_m = 198 * pfarad  # membrane capacitance
I_ex = 108 * pamp  # excitation current
g_l = 9.99 * nsiemens  # leak channel conductance
Omega_f = 3.33 * second ** -1  # rate at which docked neurotransmitters leave synapse
Omega_d = 2 * second ** -1  # rate at which the proportion of NTs ready for release grows
U_0 = 0.6  # percent docking recovery when action potential reaches synapse

N_e = 110  # Number of Excitatory Neurons
N_i = 90  # Number of Inhibitory Neurons

AD = [20, 10]  # Array Dimension

ss = 10  # electrode spacing

# Equations for Neuronal Dynamics

neuron_eqs = '''
dv/dt = (g_l *( E_l -v) + g_e *( E_e -v) + g_i *( E_i -v) +
I_ex )/ C_m : volt (unless refractory)
dg_e /dt = -g_e / tau_e : siemens # post - synaptic exc . conductance
dg_i /dt = -g_i / tau_i : siemens # post - synaptic inh . conductance
x : metre
y : metre

'''

# Group of Neurons

neurons = NeuronGroup(N_e + N_i, model=neuron_eqs,
                      threshold='v> V_th ', reset='v=V_r ',
                      refractory='tau_r ', method='euler')

neurons.v = 'E_l + rand()*( V_th -E_l )'
neurons.g_e = 'rand ()* w_e '
neurons.g_i = 'rand ()* w_i '
exc_neurons = neurons[: N_e]
inh_neurons = neurons[N_e:]

# Synapses

synapses_eqs = '''
# Usage of releasable neurotransmitter per single action potential :
du_S /dt = -Omega_f * u_S : 1 (event-driven)
# Fraction of synaptic neurotransmitter resources available :
dx_S /dt = Omega_d *(1 - x_S ) : 1 (event-driven)
'''

synapses_action = '''
u_S += U_0 * (1 - u_S )
r_S = u_S * x_S
x_S -= r_S
'''

exc_syn = Synapses(exc_neurons, neurons, model=synapses_eqs, on_pre=synapses_action + 'g_e_post += w_e *r_S ')
inh_syn = Synapses(inh_neurons, neurons, model=synapses_eqs, on_pre=synapses_action + 'g_i_post += w_i *r_S ')

MEA = np.reshape(np.random.permutation(np.array((range(N_e + N_i)))), AD)


def row(A, i):
    return np.where(A == i)[0]  # returns row of neuron on MEA


def col(A, i):
    return np.where(A == i)[1]  # Returns Column


COL = []
ROW = []

for s in range(N_e + N_i):
    COL = np.append(COL, col(MEA, s))
    ROW = np.append(ROW, row(MEA, s))

COL = COL * ss * umeter
ROW = ROW * ss * umeter

neurons.x = COL
neurons.y = ROW

exc_syn.connect(condition='i != j', p='0.8*exp(-((x_pre-x_post)**2+(y_pre-y_post)**2)/(2*(15*umeter)**2))')
inh_syn.connect(condition='i != j', p='0.4*exp(-((x_pre-x_post)**2+(y_pre-y_post)**2)/(2*(10*umeter)**2))')


def visualise_connectivity(S):
    Ns = len(S.source)
    Nt = len(S.target)
    figure(figsize=(10, 4))
    subplot(121)
    plot(zeros(Ns), arange(Ns), 'ok', ms=10)
    plot(ones(Nt), arange(Nt), 'ok', ms=10)
    for i, j in zip(S.i, S.j):
        plot([0, 1], [i, j], '-k')
    xticks([0, 1], ['Source', 'Target'])
    ylabel('Neuron index')
    xlim(-0.1, 1.1)
    ylim(-1, max(Ns, Nt))
    subplot(122)
    plot(S.i, S.j, 'ok')
    xlim(-1, Ns)
    ylim(-1, Nt)
    xlabel('Source neuron index')
    ylabel('Target neuron index')


# fig1 = visualise_connectivity(exc_syn)
# fig2 = visualise_connectivity(inh_syn)

fig3 = figure(figsize=(8, 20))
for z in range(N_e):
    plot(COL[z], ROW[z], '.r')

for i, j in zip(exc_syn.i, exc_syn.j):
    plot([COL[i], COL[j]], [ROW[i], ROW[j]], '-r')

for z in range(N_i):
    plot(COL[N_e + z], ROW[N_e + z], '.b')

for i, j in zip(inh_syn.i, inh_syn.j):
    plot([COL[i + N_e], COL[j]], [ROW[i + N_e], ROW[j]], '-b')

smon = SpikeMonitor(neurons)
stmon = StateMonitor(neurons, 'v', record=[0, 1])
run(1 * second)

fig4 = figure()
plot(smon.t / ms, smon.i, '.k')
xlabel('Time (ms)')
ylabel('Neuron index')

fig5 = figure()
subplot(211)
plot(stmon.t / ms, stmon.v[0], '-k')
xlabel('Time (ms)')
ylabel('Membrane Voltage')
subplot(212)
plot(stmon.t / ms, stmon.v[1], '-k')
xlabel('Time (ms)')
ylabel('Membrane Voltage')

count = []
for z in range(N_e + N_i):
    count = np.append(count, np.count_nonzero(smon.i == z))

data = np.zeros(AD)

for i in range(N_e + N_i):
    data[int(ROW[i]/(ss*um)), int(COL[i]/(ss*um))] = count[i]
    print(ROW[i]/ss, COL[i]/ss)

print(data)
print(count)

cmap = mpl.cm.cool
norm = mpl.colors.Normalize(vmin=0, vmax=np.max(count))

fig6, ax = plt.subplots()
ax.imshow(data, cmap=cmap, norm=norm)

show()
